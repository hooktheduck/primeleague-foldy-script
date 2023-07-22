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
all_possible_standings_data = []
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

    playoff_data = [
        playoffs_tiebreak_bool,
        playoffs_tiebreak_start,
        playoffs_tiebreak_end,
    ]
    relegation_data = [
        relegation_tiebreak_bool,
        relegation_tiebreak_start,
        relegation_tiebreak_end,
    ]
    standing_data = [standing, playoff_data, relegation_data]
    all_possible_standings_data.append(standing_data)

# FILTER TIEBREAKER FOR TEAM
if FILTER:
    for standing_data_index, standing_data in enumerate(all_possible_standings_data):
        # scan for playoff importancy:
        playoff_data = standing_data[1]

        if playoff_data[0]:
            all_possible_standings_data[standing_data_index][1][0] = False
            for n in range(playoff_data[1], playoff_data[2] + 1):
                if list(standing_data[0].keys())[n] == MYTEAM:
                    all_possible_standings_data[standing_data_index][1][0] = True

        relegation_data = standing_data[2]

        if relegation_data[0]:
            all_possible_standings_data[standing_data_index][2][0] = False
            for n in range(relegation_data[1], relegation_data[2] + 1):
                if list(standing_data[0].keys())[n] == MYTEAM:
                    all_possible_standings_data[standing_data_index][2][0] = True


workbook = openpyxl.Workbook()
sheet = workbook.active

for column, match in enumerate(MATCHES, start=1):
    sheet.cell(row=1, column=column).value = f"{match[0]} VS. {match[1]}"

for column in range(len(STANDING)):
    sheet.cell(row=1, column=column + len(MATCHES) + 2).value = column + 1

sheet.cell(row=1, column=len(STANDING) + len(MATCHES) + 3).value = f"{MYTEAM} LOCKED?"

starting_row = 2

for standing, outcome_code in zip(all_possible_standings_data, outcome_codes):
    starting_row += 1
    for match_index, char in enumerate(outcome_code):
        cell_string = OUTCOME_DICTIONARY[char]

        if OUTCOME_DICTIONARY[char].startswith("#"):
            cell_string = MATCHES[match_index][int(OUTCOME_DICTIONARY[char][1:])]

        sheet.cell(column=match_index + 1, row=starting_row).value = cell_string

        for column, team in enumerate(list(standing[0].keys()), start=len(MATCHES) + 2):
            sheet.cell(column=column, row=starting_row).value = team

            teamIndex = list(standing[0].keys()).index(team)

            playoffs_tiebreak_color = openpyxl.styles.colors.Color(
                rgb=TIEBREAK_PLAYOFFS_COLOR
            )
            playoffs_tiebreak_fill = openpyxl.styles.fills.PatternFill(
                patternType="solid", fgColor=playoffs_tiebreak_color
            )
            playoffs_tiebreak = standing[1]

            if (
                playoffs_tiebreak[0]
                and playoffs_tiebreak[1] <= teamIndex
                and playoffs_tiebreak[2] >= teamIndex
            ):
                sheet.cell(
                    column=column, row=starting_row
                ).fill = playoffs_tiebreak_fill

            relegation_tiebreak_color = openpyxl.styles.colors.Color(
                rgb=TIEBREAK_RELEGATION_COLOR
            )
            relegation_tiebreak_fill = openpyxl.styles.fills.PatternFill(
                patternType="solid", fgColor=relegation_tiebreak_color
            )
            relegation_tiebreak = standing[2]

            if (
                relegation_tiebreak[0]
                and relegation_tiebreak[1] <= teamIndex
                and relegation_tiebreak[2] >= teamIndex
            ):
                sheet.cell(
                    column=column, row=starting_row
                ).fill = relegation_tiebreak_fill

for row, standing in enumerate(all_possible_standings_data, start=3):
    playoffs_tiebreak_color = openpyxl.styles.colors.Color(rgb=TIEBREAK_PLAYOFFS_COLOR)
    playoffs_tiebreak_fill = openpyxl.styles.fills.PatternFill(
        patternType="solid", fgColor=playoffs_tiebreak_color
    )

    yes_color = openpyxl.styles.colors.Color(rgb="00cc00")
    yes_fill = openpyxl.styles.fills.PatternFill(patternType="solid", fgColor=yes_color)

    no_color = openpyxl.styles.colors.Color(rgb="ff0066")
    no_fill = openpyxl.styles.fills.PatternFill(patternType="solid", fgColor=no_color)

    lock_string = ""

    if list(standing[0].keys()).index(MYTEAM) < PLAYOFF_TEAMS:
        lock_string = "YES"
        sheet.cell(column=len(STANDING) + len(MATCHES) + 3, row=row).fill = yes_fill
    else:
        lock_string = "NO"
        sheet.cell(column=len(STANDING) + len(MATCHES) + 3, row=row).fill = no_fill

    if standing[1][0]:
        lock_string = "TIEBREAK"
        sheet.cell(
            column=len(STANDING) + len(MATCHES) + 3, row=row
        ).fill = playoffs_tiebreak_fill

    sheet.cell(column=len(STANDING) + len(MATCHES) + 3, row=row).value = lock_string


workbook.save(OUTPUT_FILE)
