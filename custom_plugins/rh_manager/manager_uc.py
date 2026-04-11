import random
from collections import defaultdict

class ManagerUC:

    # =========================
    # HEATS AUTOMÁTICOS
    # =========================

    def generate_heats_logic(pilots, pilots_per_heat, num_groups):

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
    


    # =========================
    # CATEGORÍAS
    # =========================

    def get_ranking(results):
        scores = defaultdict(int)

        for r in results:
            scores[r.pilot_id] += r.position

        ranking = sorted(scores.items(), key=lambda x: x[1])

        return [r[0] for r in ranking]