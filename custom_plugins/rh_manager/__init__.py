import logging
from flask import jsonify, request, templating
from flask.blueprints import Blueprint
from eventmanager import Evt
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

    # =========================
    # PILOTOS
    # =========================

    @bp.route('/pilots', methods=['GET'])
    def list_pilots():
        return [{
            "id": p.id,
            "name": p.name,
            "callsign": p.callsign
        } for p in rhapi.db.pilots]
    
    def pilots_map_by_callsign():
        return {
            p.callsign.strip().upper(): p
            for p in rhapi.db.pilots if p.callsign
        }
    
    @bp.route('/pilot', methods=['POST'])
    def add_edit_pilot():
        data = request.get_json()

        name = data.get("name")
        callsign = data.get("callsign")
        callsign = callsign.strip().upper()

        if not callsign:
            return {"status": "error", "msg": "alias requerido"}, 400


        pilots_map = pilots_map_by_callsign()
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
    # HEATS AUTOMÁTICOS
    # =========================

    def create_heats_by_groups(groups):
        pilots_map = pilots_map_by_callsign()
        for g_idx, heats in enumerate(groups):
            for h_idx, heat in enumerate(heats):

                # TODO - VERIFICAR SI NO ES NECESARIO CREAR UN NUEVO GRUPO (raceclass)
                heat_id = rhapi.db.heat_add(
                    name=f"G{g_idx+1}-H{h_idx+1}"
                )

                for slot, cs in enumerate(heat):
                    p = pilots_map.get(cs)
                    if p:
                        # TODO - VERIFICAR COMO SE OBTIENEN LOS SLOTS PARA EDITARLOS
                        # AL PARECER NO HAY FORMA DE AGREGAR NUEVOS
                        rhapi.db.heat_slot_add(
                            heat_id=heat_id,
                            node_index=slot,
                            pilot_id=p.id
                        )

    @bp.route('/heats/generate', methods=['POST'])
    def generate_heats():
        data = request.get_json()
        pilots_per_heat = data["pilots_per_heat"]
        num_groups = data["num_groups"]

        # CAMBIAR ALGORITMO PARA USAR LOS IDS Y NO EL CALLSIGN
        pilots = [p.callsign for p in rhapi.db.pilots]
        groups = muc.generate_heats_logic(pilots, pilots_per_heat, num_groups)

        # limpiar heats existentes
        for h in rhapi.db.heats:
            rhapi.db.heat_delete(h.id)

        # crear nuevos
        ## TODO - MODIFICAR PARA OBTENER LOS IDS DE CADA HEAT Y SLOTS
        create_heats_by_groups(groups)

        return {"status": "ok", "groups": groups}

    @bp.route('/heat/<id>', methods=['POST'])
    def edit_heat(id):
        data = request.get_json()
        heat = rhapi.db.heat_by_id(id)
        if heat:
            slots = {
                s.id: s
                for s in rhapi.db.slots_by_heat(heat.id)
            }
            slotsForUpdate = [{
                "slot_id": s.id,
                "pilot": s.pilot_id
            } for s in data if slots.get(s.id)]

            rhapi.db.slots_alter_fast(slotsForUpdate)
        else:
            return {"status": "error", "msg": "not found"}, 400

        return {"status": "ok"}

    @bp.route('/raceclass/remix/<id>', methods=['post'])
    def remix_class(id):

        raceClass = rhapi.db.raceclass_by_id(id)
        if raceClass:
            heats_by_class = rhapi.db.heats_by_class(raceClass.id)
            slotsByHeats = {
                r.id: rhapi.db.slots_by_heat(r.id)
                for r in heats_by_class
            }
            pilots = {
                s.id: rhapi.db.pilot_by_id(s.pilot_id)
                for s in slotsByHeats
            }
            first_slot = next(iter(slotsByHeats.values()))
            pilots_per_heat = len(vars(first_slot))
            num_groups = len(heats_by_class)

            for h in rhapi.db.heats_by_class(raceClass.id):
                rhapi.db.heat_delete(h.id)

            groups = muc.generate_heats_logic(pilots, pilots_per_heat, num_groups)
            ## TODO - MODIFICAR PARA OBTENER LOS IDS DE CADA HEAT Y SLOTS
            create_heats_by_groups(groups)
        else:
            return {"status": "error", "msg": "not found"}, 400

        return {"status": "ok", "groups": groups}

    @bp.route('/raceclass/<id>', methods=['DELETE'])
    def delete_class(id):

        raceClass = rhapi.db.raceclass_by_id(id)
        if raceClass:
            for h in rhapi.db.heats_by_class(raceClass.id):
                rhapi.db.heat_delete(h.id)

            rhapi.db.raceclass_delete(raceClass.id)
        else:
            return {"status": "error", "msg": "not found"}, 400

        return {"status": "ok"}


    @bp.route('/heats', methods=['DELETE'])
    def delete_all_heats():
        for h in rhapi.db.heats:
            rhapi.db.heat_delete(h.id)

        return {"status": "ok"}

    # =========================
    # CATEGORÍAS
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
    # FINALES
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
        return jsonify({"status": "running"})

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
