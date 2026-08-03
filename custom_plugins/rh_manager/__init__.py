import json
import logging
from flask import jsonify, request
from flask.blueprints import Blueprint
from .alchemy_encoder import AlchemyEncoder
from .manager_uc import ManagerUC

logger = logging.getLogger(__name__)
muc = ManagerUC()

def initialize(rhapi):

    logger.info("[RH Manager] Plugin initialized")

    bp = Blueprint(
        'rh_manager',
        __name__,
        url_prefix='/api/rhm'
    )

    def json_response(data, code = 201):
        return json.dumps(data, cls=AlchemyEncoder), code, {'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*'}

    # =========================
    # PILOTS
    # =========================

    @bp.route('/pilots', methods=['GET'])
    def list_pilots():
        return [{
            "id": p.id,
            "name": p.name,
            "callsign": p.callsign
        } for p in rhapi.db.pilots]
    
    @bp.route('/pilot', methods=['POST'])
    def add_edit_pilot():
        data = request.get_json()

        name = data.get("name")
        callsign = data.get("callsign")
        callsign = callsign.strip().upper()

        if not callsign:
            return {"status": "error", "msg": "alias requerido"}, 400


        pilots_map = muc.pilots_map_by_callsign(rhapi)
        existing = pilots_map.get(callsign)
        if existing and name:
            rhapi.db.pilot_alter(existing.id, name=name)

            return {
                "status": "exists",
                "pilot": {
                    "id": existing.id,
                    "name": existing.name,
                    "callsign": existing.callsign
                }
            }

        # ✔ crear si no existe
        new_pilot = rhapi.db.pilot_add(
            name=name,
            callsign=callsign
        )

        return {
            "status": "created",
            "pilot": {
                "id": new_pilot.id,
                "name": name,
                "callsign": callsign
            }
        }
    
    @bp.route('/pilot/<id>', methods=['DELETE'])
    def delete_pilot(id):
        existing = rhapi.db.pilot_by_id(id)
        if existing:
            rhapi.db.pilot_delete(existing.id)

        return {"status": "ok"}
    
    # =========================
    # CONFIG
    # =========================

    @bp.route('/frequencies', methods=['GET'])
    def get_frequencies():    
        return jsonify({"status": "ok", "frequencies": muc.get_frequencies(rhapi)})

    @bp.route('/race_config_data', methods=['GET'])
    def race_config_data():
        data = {
            'formats': muc.get_formats(rhapi),
            'win_conditions': muc.get_win_conditions(rhapi),
            'frequencies': muc.get_frequencies(rhapi),
        }

        return json_response({"status": "ok", "data": data})
    
    # =========================
    # AUTOMATICS HEATS
    # =========================

    @bp.route('/heats/generate', methods=['POST'])
    def generate_heats():
        try:
            data = request.get_json()

            # always-required fields, regardless of new-class vs existing-class
            if (not data
                or 'num_groups' not in data
                or 'pilots_per_heat' not in data
                or 'frequencies' not in data
                or 'pilots' not in data
                or 'delete_everything' not in data):
                return json_response({"error": "Faltan campos requeridos"}, 400)

            existing_class = None
            if data.get('class_id'):
                existing_class = rhapi.db.raceclass_by_id(int(data['class_id']))
                if not existing_class:
                    return json_response({"error": "La clase existente indicada no existe"}, 400)
            else:
                # only required when creating a brand new class -- an
                # existing class already has its own format/win_condition/rounds
                if ('stage' not in data or 'format' not in data
                        or 'rounds_per_group' not in data or 'win_condition' not in data):
                    return json_response({"error": "Faltan campos requeridos"}, 400)
                data["format"] = int(data["format"])
                data["rounds_per_group"] = int(data["rounds_per_group"])

            data["num_groups"] = int(data["num_groups"])
            data["pilots_per_heat"] = int(data["pilots_per_heat"])
            data["frequencies"] = [int(f) for f in data["frequencies"].split(",")]
            data["pilots"] = [int(p) for p in data["pilots"].split(",")]
            data["delete_everything"] = bool(data["delete_everything"])
            if data["pilots_per_heat"] > len(data["frequencies"]):
                return json_response({"error": "Las frecuencias y la cantidad de pilotos debe coincidir"}, 400)

            pilots_ids = muc.pilots_map_by_id(rhapi)
            pilots_selected = [pilots_ids.get(p) for p in data["pilots"]]
            pilots = [p.callsign for p in pilots_selected]
            groups = muc.generate_heats_logic(pilots, data["pilots_per_heat"], data["num_groups"])

            if data["delete_everything"]:
                rhapi._racecontext.race.reset_current_laps()
                rhapi._racecontext.rhdata.clear_race_data()
                rhapi._racecontext.rhdata.reset_heats()
                rhapi._racecontext.rhdata.reset_raceClasses()
                existing_class = None  # it no longer exists after reset_raceClasses()

            # crear nuevos (o agregar a la clase existente)
            created = muc.create_heats_by_groups(rhapi, data, groups, existing_class=existing_class)

            # build the response the same way as /raceclass/<id>/results and
            # remix (freq_label, win_condition, standings...) instead of the
            # bare create_heats_by_groups/map_race_data shape, which was
            # missing those fields on first generation
            race_classes = {}
            for rc in created['classes']:
                grid = muc.get_class_results_grid(rhapi, rc.id)
                if grid:
                    race_classes.update(grid)

        except (ValueError, TypeError) as e:
            return json_response({"error": str(e)}, 422)

        return json_response({"status": "ok", "groups": race_classes})

    @bp.route('/heat/<id>', methods=['POST'])
    def edit_heat(id):
        data = request.get_json()
        heat = rhapi.db.heat_by_id(id)
        if not heat:
            return json_response({"status": "error", "msg": "not found"}, 400)

        # `data` is a list of {"slot_id": int, "pilot": int} sent by the client
        slots = {
            s.id: s
            for s in rhapi.db.slots_by_heat(heat.id)
        }
        slotsForUpdate = [{
            "slot_id": s.get("slot_id"),
            "pilot": s.get("pilot")
        } for s in data if slots.get(s.get("slot_id"))]

        rhapi.db.slots_alter_fast(slotsForUpdate)

        return json_response({"status": "ok"})

    @bp.route('/raceclass/remix/<id>', methods=['POST'])
    def remix_class(id):

        raceClass = rhapi.db.raceclass_by_id(id)
        if not raceClass:
            return json_response({"status": "error", "msg": "not found"}, 400)

        heats_by_class = rhapi.db.heats_by_class(raceClass.id)
        if not heats_by_class:
            return json_response({"status": "error", "msg": "class has no heats to remix"}, 400)

        # optional body: {"group_ids": [0, 2]} -- remix only those groups
        # (by their group_id number, 0-based) instead of the whole class
        data = request.get_json(silent=True) or {}
        group_ids = data.get("group_ids")
        if group_ids is not None:
            group_ids = [int(g) for g in group_ids]

        ok, err_msg = muc.remix_raceclass_heats(rhapi, raceClass, group_ids=group_ids)
        if not ok:
            return json_response({"status": "error", "msg": err_msg or "remix failed"}, 400)

        grid = muc.get_class_results_grid(rhapi, raceClass.id)

        return json_response({"status": "ok", "groups": grid})

    @bp.route('/raceclass/<id>/groups', methods=['GET'])
    def raceclass_groups(id):
        raceClass = rhapi.db.raceclass_by_id(id)
        if not raceClass:
            return json_response({"status": "error", "msg": "not found"}, 400)

        return json_response({"status": "ok", "groups": muc.get_class_groups(rhapi, int(id))})

    @bp.route('/raceclasses/results', methods=['GET'])
    def all_raceclasses_results():
        grid = muc.get_all_classes_results_grid(rhapi)
        return json_response({"status": "ok", "groups": grid})

    @bp.route('/raceclasses', methods=['GET'])
    def list_raceclasses():
        return json_response({"status": "ok", "classes": muc.get_raceclasses_summary(rhapi)})

    @bp.route('/raceclass/<id>/results', methods=['GET'])
    def raceclass_results(id):
        grid = muc.get_class_results_grid(rhapi, int(id))
        if grid is None:
            return json_response({"status": "error", "msg": "not found"}, 400)

        return json_response({"status": "ok", "groups": grid})

    @bp.route('/heat/<id>/results', methods=['GET'])
    def heat_results(id):
        grid = muc.get_heat_results_grid(rhapi, int(id))
        if grid is None:
            return json_response({"status": "error", "msg": "not found"}, 400)

        return json_response({"status": "ok", "heat": grid})

    @bp.route('/raceclass/<id>', methods=['DELETE'])
    def delete_class(id):

        raceClass = rhapi.db.raceclass_by_id(id)
        if not raceClass:
            return json_response({"status": "error", "msg": "not found"}, 400)

        ok, err_msg = muc.delete_class_safe(rhapi, raceClass)
        if not ok:
            return json_response({"status": "error", "msg": err_msg}, 400)

        return {"status": "ok"}


    @bp.route('/heats', methods=['DELETE'])
    def delete_all_heats():
        # heat_delete silently refuses (returns False, no exception) for a
        # heat that still has a saved race attached -- report which ones
        # instead of claiming "ok" when some heats are actually still there
        blocked = [h.display_name for h in rhapi.db.heats if not rhapi.db.heat_delete(h.id)]
        if blocked:
            return json_response({
                "status": "error",
                "msg": "No se pudieron eliminar (tienen carreras guardadas): " + ", ".join(blocked)
            }, 400)

        return {"status": "ok"}

    @bp.route('/heat/<id>', methods=['DELETE'])
    def delete_heat(id):
        heat = rhapi.db.heat_by_id(id)
        if not heat:
            return json_response({"status": "error", "msg": "not found"}, 400)

        if not rhapi.db.heat_delete(heat.id):
            return json_response({"status": "error", "msg": f"El heat '{heat.display_name}' tiene una carrera guardada y no se puede eliminar"}, 400)

        return json_response({"status": "ok"})

    # =========================
    # CATEGORIES
    # =========================

    @bp.route('/categories/generate', methods=['POST'])
    def generate_categories():

        data = request.get_json()
        splits = data["splits"]  # ej: [8, 999]

        ranking = muc.get_ranking(rhapi.db.race_results())
        categories = []
        start = 0
        for s in splits:
            cat = ranking[start:start+s]
            categories.append(cat)
            start += s

        return {"categories": categories}

    # =========================
    # FINALS
    # =========================

    @bp.route('/finals/generate', methods=['POST'])
    def generate_finals():

        data = request.get_json()
        top_n = data.get("top", 8)

        ranking = muc.get_ranking(rhapi.db.race_results())
        top = ranking[:top_n]

        # distribución
        semiA = [top[0], top[2], top[4], top[6]]
        semiB = [top[1], top[3], top[5], top[7]]

        return {
            "semiA": semiA,
            "semiB": semiB
        }


    # =========================
    # STATUS
    # =========================

    @bp.route('/status', methods=['GET'])
    def status():
        return json_response({"status": "running", "server_ip": muc.get_server_ip()})

    # =========================
    # UI
    # =========================

    @bp.route('/ui')
    def ui():
        return """
        <html>
        <head>
            <title>Race Manager</title>
        </head>
        <body>

        <h2>RotorHazard Manager</h2>

        <button onclick="generateHeats()">Generar Heats</button>
        <button onclick="generateBracket()">Doble Eliminación</button>

        <script>
        function generateHeats() {
            fetch('/rhapi/heats/generate', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    pilots_per_heat: 4,
                    num_groups: 3
                })
            })
        }

        function generateBracket() {
            fetch('/rhapi/bracket/double_elimination', {
                method: 'POST'
            })
        }
        </script>

        </body>
        </html>
        """

    # ============================================
    # REGISTER ENDPOINTS
    # ============================================

    rhapi.ui.blueprint_add(bp)
