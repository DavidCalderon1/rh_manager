import json
import logging
import random
import socket
from collections import defaultdict

logger = logging.getLogger(__name__)

class ManagerUC:

    def get_server_ip(self):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(('8.8.8.8', 1))
            local_ip = s.getsockname()[0]
        except Exception:
            local_ip = '127.0.0.1'
        finally:
            s.close()

        return local_ip
    # =========================
    # MAPPERS
    # =========================
    
    def pilots_map_by_id(self, rhapi):
        return {
            p.id: p
            for p in rhapi.db.pilots if p.callsign
        }
    
    def pilots_map_by_callsign(self, rhapi):
        return {
            p.callsign.strip().upper(): p
            for p in rhapi.db.pilots if p.callsign
        }
    
    def map_race_data(self, data):
        race_data = {c['id']: {**c, "heats": {}} for c in data['classes']}

        for h in data['heats']:
            c_id = h['class_id']
            if c_id in race_data:
                h_id = h['id']
                h['nodes'] = [] 
                race_data[c_id]["heats"][h_id] = h

        for n in data['nodes']:
            h_id = n['heat_id']
            for c_id in race_data:
                if h_id in race_data[c_id]["heats"]:
                    race_data[c_id]["heats"][h_id]["nodes"].append(n)
                    break

        return race_data

    def get_raceclasses_summary(self, rhapi):
        """Lightweight id+name listing of every race class -- for populating
        an "existing class" picker in Excel without paying for the full
        results grid (get_all_classes_results_grid)."""
        return [{"id": rc.id, "name": rc.display_name} for rc in rhapi.db.raceclasses]

    def get_class_groups(self, rhapi, class_id):
        """Heat-groups (G1, G2, ...) within a class and which heats belong to
        each -- lets a caller (e.g. the remix picker in Excel) target a
        subset of groups instead of the whole class."""
        heats = rhapi.db.heats_by_class(class_id)
        by_group = defaultdict(list)
        for h in heats:
            gid = h.group_id if h.group_id is not None else 0
            by_group[gid].append({"id": h.id, "display_name": h.display_name})

        return [
            {"group_id": gid, "label": f"G{gid + 1}", "heats": group_heats}
            for gid, group_heats in sorted(by_group.items())
        ]

    def get_node_freq_label(self, rhapi, node_index):
        """Real band+channel label for a physical node index, from the active
        frequency profile -- not to be confused with the per-race generation
        `frequencies` list, which is just a list of *available* slots."""
        if node_index is None:
            return None
        currentProfileID = rhapi._racecontext.rhdata.get_option('currentProfile')
        profile = rhapi.db.frequencyset_by_id(currentProfileID)
        profile_frequencies = json.loads(profile.frequencies)
        bands = profile_frequencies.get('b', [])
        channels = profile_frequencies.get('c', [])
        if node_index >= len(bands) or node_index >= len(channels):
            return None
        band = bands[node_index]
        channel = channels[node_index]
        if not band or not channel:
            return None
        return f"{band}{channel}"

    def build_heat_grid(self, rhapi, heat, class_id=None):
        """Shared builder: one heat's nodes, enriched with slot_id, the pilot's
        real frequency label, and a `rounds` array per node with laps/time/
        position/points already recorded for each round run so far (rounds not
        yet run are simply absent -- zero-filling is a rendering concern)."""

        class_id = class_id if class_id is not None else heat.class_id
        slots = rhapi.db.slots_by_heat(heat.id)

        # a slot can have node_index=None -- a pilot seeded into the heat
        # but not yet resolved to a physical node/frequency (RotorHazard
        # shows this as a "Seed Now" prompt, typically when there are fewer
        # configured frequencies than pilots needing one). Race results are
        # always reported per physical node, so a slot with no node can
        # never be matched to one anyway -- keep those separate instead of
        # keying the lookup dict by node_index, which would silently drop
        # every unassigned slot but the last (None is not a unique key).
        nodes_by_index = {}
        unassigned_nodes = []
        for s in slots:
            node = {
                "slot_id": s.id,
                "pilot_id": s.pilot_id,
                "node_index": s.node_index,
                "freq_label": self.get_node_freq_label(rhapi, s.node_index),
                "rounds": [],
            }
            if s.node_index is None:
                unassigned_nodes.append(node)
            else:
                nodes_by_index[s.node_index] = node

        for race in rhapi.db.races_by_heat(heat.id):
            results = rhapi.db.race_results(race.id)
            if not results:
                continue
            primary = results.get("meta", {}).get("primary_leaderboard")
            for row in results.get(primary, []):
                node = nodes_by_index.get(row.get("node"))
                if node is not None:
                    node["rounds"].append({
                        "round": race.round_id,
                        "laps": row.get("laps"),
                        "time": row.get("total_time"),
                        "position": row.get("position"),
                        "points": row.get("points"),
                    })

        return {
            "id": heat.id,
            "display_name": heat.display_name,
            "class_id": class_id,
            "nodes": list(nodes_by_index.values()) + unassigned_nodes,
        }

    def get_class_standings(self, rhapi, class_id):
        """Overall class standings (position across all rounds so far),
        using the class's own configured ranking method -- same data source
        MultiGP/FPVScores sync uses (rhapi.db.raceclass_ranking)."""

        rankings = rhapi.db.raceclass_ranking(class_id)
        if not rankings:
            return None

        meta = rankings.get("meta", {}) or {}
        standings = []
        for rank in rankings.get("ranking", []):
            row = dict(rank)
            pilot_id = row.pop("pilot_id", None)
            callsign = row.pop("callsign", None)
            position = row.pop("position", None)
            team_name = row.pop("team_name", None)
            row.pop("node", None)
            standings.append({
                "pilot_id": pilot_id,
                "callsign": callsign,
                "position": position,
                "team_name": team_name,
                "extra": row,
            })

        return {
            "method_label": meta.get("method_label"),
            "rank_fields": meta.get("rank_fields"),
            "standings": standings,
        }

    def get_class_results_grid(self, rhapi, class_id):
        """groups -> heats -> nodes shape (same as create_heats_by_groups/
        map_race_data), enriched via build_heat_grid with real results/freq,
        plus overall class standings."""

        raceClass = rhapi.db.raceclass_by_id(class_id)
        if not raceClass:
            return None

        heats_out = {
            heat.id: self.build_heat_grid(rhapi, heat, class_id)
            for heat in rhapi.db.heats_by_class(class_id)
        }

        return {
            raceClass.id: {
                "id": raceClass.id,
                "display_name": raceClass.display_name,
                "rounds": raceClass.rounds,
                "win_condition": raceClass.win_condition,
                "heats": heats_out,
                "standings": self.get_class_standings(rhapi, class_id),
            }
        }

    def get_all_classes_results_grid(self, rhapi):
        """Every race class currently in RotorHazard, not just the ones the
        Excel workbook happens to remember generating -- lets the workbook
        pick up classes/heats created directly in the RotorHazard UI, via
        MultiGP import, etc."""

        merged = {}
        for raceClass in rhapi.db.raceclasses:
            grid = self.get_class_results_grid(rhapi, raceClass.id)
            if grid:
                merged.update(grid)
        return merged

    def get_heat_results_grid(self, rhapi, heat_id):
        """Single-heat variant used by the lightweight per-heat refresh button --
        same node/round shape as get_class_results_grid but scoped to one heat,
        avoiding a full-class fetch for a single-heat refresh."""

        heat = rhapi.db.heat_by_id(heat_id)
        if not heat:
            return None

        raceClass = rhapi.db.raceclass_by_id(heat.class_id) if heat.class_id else None

        heat_grid = self.build_heat_grid(rhapi, heat)
        heat_grid["win_condition"] = raceClass.win_condition if raceClass else ""
        return heat_grid

    # =========================
    # CONFIG
    # =========================

    def get_formats(self, rhapi):
        return [{
            'id': rf.id,
            'name': rf.name,
        } for rf in rhapi.db.raceformats]

    def get_win_conditions(self, rhapi):
        return [{
            'name': name,
            'label': method.label,
        } for name, method in rhapi.classrank.methods.items()]

    def get_frequencies(self, rhapi):
        currentProfileID = rhapi._racecontext.rhdata.get_option('currentProfile')
        profile = rhapi.db.frequencyset_by_id(currentProfileID)
        profile_frequencies = json.loads(profile.frequencies)
        return [{
            'id': idx,
            'frecuency': frq,
            'name': '{0}{1}'.format(profile_frequencies['b'][idx], profile_frequencies['c'][idx])
        } for idx, frq in enumerate(profile_frequencies['f']) if frq > 0]

    # =========================
    # HEATS AUTOMATICS
    # =========================

    def generate_heats_logic(self, pilots, pilots_per_heat, num_groups):

        encounters = defaultdict(lambda: defaultdict(int))
        all_groups = []

        for g in range(num_groups):

            offset = (g * 3) % len(pilots)
            rotated = pilots[offset:] + pilots[:offset]

            heats = []
            for i in range(0, len(rotated), pilots_per_heat):
                heat = rotated[i:i+pilots_per_heat]

                best = heat[:]
                best_cost = 999999

                for _ in range(15):
                    random.shuffle(heat)
                    cost = sum(
                        encounters[a][b]
                        for i, a in enumerate(heat)
                        for b in heat[i+1:]
                    )
                    if cost < best_cost:
                        best = heat[:]
                        best_cost = cost

                heats.append(best)

                for i in range(len(best)):
                    for j in range(i+1, len(best)):
                        encounters[best[i]][best[j]] += 1
                        encounters[best[j]][best[i]] += 1

            all_groups.append(heats)

        return all_groups
    
    def create_heats_by_groups(self, rhapi, data, groups, existing_class=None):
        """Creates every group as heats within a SINGLE race class -- either a
        brand new one (named after data['stage']), or an existing class the
        caller passed in (existing_class), in which case new groups continue
        numbering after whatever group_id's are already used in that class
        (so "adding heats to an existing class" doesn't collide with the
        groups it already has)."""

        pilots_map = self.pilots_map_by_callsign(rhapi)

        if existing_class is not None:
            raceclass = existing_class
            existing_heats = rhapi.db.heats_by_class(raceclass.id)
            used_group_ids = [h.group_id for h in existing_heats if h.group_id is not None]
            group_offset = (max(used_group_ids) + 1) if used_group_ids else 0
        else:
            raceclass = rhapi.db.raceclass_add(
                name=data['stage'],
                raceformat=data['format'],
                win_condition=data['win_condition'],
                rounds=data['rounds_per_group'],
                heat_advance_type=1,
            )
            group_offset = 0

        class_heats = []
        heat_nodes = []
        for g_idx, heats in enumerate(groups):
            group_num = group_offset + g_idx
            for h_idx, heat in enumerate(heats):
                initPilots = {}
                for slot, cs in enumerate(heat):
                    p = pilots_map.get(cs.strip().upper())
                    if p:
                        initPilots[data['frequencies'][slot]] = p.id
                heat = rhapi._racecontext.rhdata.add_heat(
                    init={
                        'name': f"{raceclass.name} G{group_num + 1}-H{h_idx + 1}",
                        'class_id': raceclass.id,
                        'group_id': group_num,
                    },
                    initPilots=initPilots
                )
                class_heats.append(heat)
                if len(heat_nodes):
                    heat_nodes.extend(rhapi.db.slots_by_heat(heat.id))
                else:
                    heat_nodes = rhapi.db.slots_by_heat(heat.id)

        return {'classes': [raceclass], 'heats': class_heats, 'nodes': heat_nodes}

    def remix_raceclass_heats(self, rhapi, raceClass, group_ids=None):
        """Reshuffle the pilots of one or more heat-groups within a race
        class, using RotorHazard's own "Balanced Random Fill (Minimize
        Repeats)" heat generator (data/plugins/heatgen_minimal_repeats,
        generator name 'minimal_repeat_fill') instead of a home-grown
        shuffle. Each target group is remixed INDEPENDENTLY -- its own heats
        deleted and regenerated from its own occupied pilots only -- so
        groups never touched here (and pilots in them) are left exactly as
        they were, and pilots never move between groups. group_ids=None
        remixes every group in the class (the original whole-class
        behavior, still used by the class-level "MX" button when the caller
        doesn't scope it down).

        Returns (ok, error_msg) instead of a bare bool -- RotorHazard's
        heat_delete silently refuses (returns False, no exception) when a
        heat still has a saved race attached, and a caller that doesn't
        check for that ends up deleting SOME of a group's heats and then
        generating new ones alongside the ones that refused to go, leaving
        duplicates. Checking every heat in a group up front, before
        deleting any of them, avoids that half-done state."""

        all_heats = rhapi.db.heats_by_class(raceClass.id)
        if not all_heats:
            return False, "La clase no tiene heats para remixar"

        heats_by_group = defaultdict(list)
        for h in all_heats:
            gid = h.group_id if h.group_id is not None else 0
            heats_by_group[gid].append(h)

        target_group_ids = list(group_ids) if group_ids is not None else list(heats_by_group.keys())
        target_group_ids = [g for g in target_group_ids if g in heats_by_group]
        if not target_group_ids:
            return False, "No se encontraron los grupos indicados en esta clase"

        for gid in target_group_ids:
            group_heats = heats_by_group[gid]

            blocked = [h for h in group_heats if rhapi._racecontext.rhdata.savedRaceMetas_has_heat(h.id)]
            if blocked:
                names = ', '.join(h.display_name for h in blocked)
                return False, f"El grupo G{gid + 1} tiene heats con carreras guardadas ({names}) y no se puede remixar"

            pilot_ids = []
            pilots_per_heat = 0
            for heat in group_heats:
                slots = rhapi.db.slots_by_heat(heat.id)
                occupied = [s for s in slots if s.pilot_id]
                pilots_per_heat = max(pilots_per_heat, len(occupied))
                for slot in occupied:
                    pilot_ids.append(slot.pilot_id)

            if not pilot_ids:
                continue

            ids_before = {h.id for h in rhapi.db.heats_by_class(raceClass.id)}
            for heat in group_heats:
                rhapi.db.heat_delete(heat.id)

            result = rhapi._racecontext.heat_generate_manager.generate('minimal_repeat_fill', {
                'output_class': raceClass.id,
                'qualifiers_per_heat': pilots_per_heat,
                'pilot_ids': ','.join(str(p) for p in pilot_ids),
            })
            if result is False:
                return False, f"El generador nativo fallo al remixar el grupo G{gid + 1}"

            # the native generator creates brand-new heats with generic
            # names and no group_id of its own -- reassign them back to
            # THIS group's id and this class's G<n>-H<n> naming convention
            # (rhapi.heat_alter doesn't expose group_id, so this goes
            # through rhdata directly, same as create_heats_by_groups does
            # for add_heat).
            new_heats = sorted(
                (h for h in rhapi.db.heats_by_class(raceClass.id) if h.id not in ids_before),
                key=lambda h: h.id
            )
            for idx, heat in enumerate(new_heats):
                rhapi._racecontext.rhdata.alter_heat({
                    'heat': heat.id,
                    'name': f"{raceClass.name} G{gid + 1}-H{idx + 1}",
                    'group_id': gid,
                })

        return True, None

    def delete_class_safe(self, rhapi, raceClass):
        """Delete a class and its heats. Two things a naive
        heat_delete-then-raceclass_delete sequence gets wrong:
        (1) both calls silently return False (no exception) when the heat
        or class still has a saved race attached -- RotorHazard protects
        race history from being orphaned -- so a caller that doesn't check
        the return value ends up reporting success when nothing was
        actually deleted;
        (2) even when nothing blocks it, raceclass_delete can still hit a
        raw "FOREIGN KEY constraint failed" error deleting the race_class
        row itself -- observed live against this container on a fresh test
        class with zero saved races -- because its RaceClassAttribute rows
        don't reliably get flushed to the DB before the race_class row in
        the same SQLAlchemy commit (there's no declared ORM relationship
        between them for SQLAlchemy to order by). Deleting those attribute
        rows here, in their own commit, before calling raceclass_delete
        sidesteps that ordering issue entirely."""
        import Database

        for h in rhapi.db.heats_by_class(raceClass.id):
            if not rhapi.db.heat_delete(h.id):
                return False, f"El heat '{h.display_name}' tiene una carrera guardada y no se puede eliminar"

        for attr in rhapi.db.raceclass_attributes(raceClass.id):
            Database.DB_session.delete(attr)
        Database.DB_session.commit()

        if not rhapi.db.raceclass_delete(raceClass.id):
            return False, "La clase tiene una carrera guardada y no se puede eliminar"

        return True, None

    # =========================
    # CATEGORIES
    # =========================

    def get_ranking(self, results):
        scores = defaultdict(int)

        for r in results:
            scores[r.pilot_id] += r.position

        ranking = sorted(scores.items(), key=lambda x: x[1])

        return [r[0] for r in ranking]