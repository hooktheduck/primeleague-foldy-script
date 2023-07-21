# imports
import numpy as np
import copy
import openpyxl


# CONSTANTS
MATCHES = [["A", "B"], ["C", "D"]]
DELTA_SCORE = [[[2, 0], [0, 2]], [[0, 2], [2, 0]], [[1, 1], [1, 1]]]
OUTCOME_DICTIONARY = {"0": "#0", "1": "#1", "2": "DRAW"}
STANDING = {"A": [0, 0], "B": [0, 0], "C": [0, 0], "D": [0, 0]}
FILTER = True
MYTEAM = "A"
PLAYOFF_TEAMS = 2
RELEGATION_TEAMS = 2
TIEBREAK_PLAYOFFS_COLOR = "3399ff"
TIEBREAK_RELEGATION_COLOR = "ff5050"
OUTPUT_FILE = "test.xlsx"


# GENERATE ALL POSSIBLE OUTCOME CODES
outcome_codes = []
for code in range(len(DELTA_SCORE) ** len(MATCHES)):
    outcome_codes.append(
        str.zfill(np.base_repr(code, base=len(DELTA_SCORE)), len(MATCHES))
    )


# USE CODES AND DELTA_SCORE TO GENERATE STANDINGS
all_possible_standings = []
for code in outcome_codes:
    delta_standing = copy.deepcopy(STANDING)

    for matchIndex, matchCode in enumerate(code):
        team_A = MATCHES[matchIndex][0]
        team_B = MATCHES[matchIndex][1]
        delta_A = DELTA_SCORE[int(matchCode)][0]
        delta_B = DELTA_SCORE[int(matchCode)][1]

        delta_standing[team_A][0] += delta_A[0]
        delta_standing[team_A][1] += delta_A[1]
        delta_standing[team_B][0] += delta_B[0]
        delta_standing[team_B][1] += delta_B[1]

    delta_standing_sorted = dict(
        sorted(delta_standing.items(), key=lambda x: x[1][0], reverse=True)
    )
    all_possible_standings.append(delta_standing_sorted)


# GENERATE TIEBREAKER DATA
tiebreaker = []
for standing in all_possible_standings:
    score_list = list(standing.values())
    playoffs_tiebreak_start = 0
    playoffs_tiebreak_end = PLAYOFF_TEAMS
    playoffs_tiebreak_bool = False

    relegation_tiebreak_start = 0
    relegation_tiebreak_end = len(score_list) - RELEGATION_TEAMS
    relegation_tiebreak_bool = False

    # CHECK FOR PLAYOFF TIEBREAKER
    if score_list[PLAYOFF_TEAMS] == score_list[PLAYOFF_TEAMS - 1]:
        playoffs_tiebreak_bool = True
        for score in score_list:
            if score == score_list[PLAYOFF_TEAMS]:
                break
            playoffs_tiebreak_start += 1

        for n in range(PLAYOFF_TEAMS, len(score_list)):
            if score_list[n] != score_list[PLAYOFF_TEAMS]:
                break
            playoffs_tiebreak_end += 1

        playoffs_tiebreak_end -= 1

    # CHECK FOR RELEGATION TIEBREAKER
    if (
        score_list[len(score_list) - RELEGATION_TEAMS - 1]
        == score_list[len(score_list) - RELEGATION_TEAMS]
    ):
        relegation_tiebreak_bool = True

        for score in score_list:
            if score == score_list[len(score_list) - RELEGATION_TEAMS]:
                break
            relegation_tiebreak_start += 1

        for n in range(len(score_list) - RELEGATION_TEAMS, len(score_list)):
            if score_list[n] != score_list[len(score_list) - RELEGATION_TEAMS]:
                break
            relegation_tiebreak_end += 1

        relegation_tiebreak_end -= 1

    tiebreakerData = [
        [playoffs_tiebreak_bool, playoffs_tiebreak_start, playoffs_tiebreak_end],
        [relegation_tiebreak_bool, relegation_tiebreak_start, relegation_tiebreak_end],
    ]

    tiebreaker.append(tiebreakerData)

# FILTER TIEBREAKER FOR TEAM
if FILTER:
    for standing_data_index, standing_data in enumerate(all_possible_standings):
        playoff_data = tiebreaker[standing_data_index][0]
        relegation_data = tiebreaker[standing_data_index][1]

        # CHECK IF TEAM IN PLAYOFF TIEBREAKER IS MYTEAM
        if playoff_data[0]:
            playoff_data[0] = False
            for n in range(playoff_data[1], playoff_data[2] + 1):
                if list(standing_data.keys())[n] == MYTEAM:
                    playoff_data[0] = True

        # CHECK IF TEAM IN RELEGATION TIEBREAKER IS MYTEAM
        if relegation_data[0]:
            relegation_data[0] = False
            for n in range(relegation_data[1], relegation_data[2] + 1):
                if list(standing_data.keys())[n] == MYTEAM:
                    relegation_data[0] = True
