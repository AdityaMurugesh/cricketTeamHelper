import pandas as pd
import numpy as np

file = 'data/matchdata.xlsx'

#Reading Data

df_match_info = pd.read_excel(file, sheet_name=0)
df_batting = pd.read_excel(file, sheet_name=1)
df_bowling = pd.read_excel(file, sheet_name=2)

#Calculating Bowler Figures

figures = (
    df_bowling.groupby(['match_id', 'bowler'], as_index=False).agg
    (
        overs = ('over', 'count'),
        runs = ('runs_conceded', 'sum'),
        wickets = ('wickets', 'sum'),
        extras = ('extras', 'sum'),
        fours = ('fours', 'sum'),
        sixes = ('sixes', 'sum'),
    )
        )

figures['economy'] = (figures['runs'] / figures['overs']).round(2)
figures['strike_rate'] = figures['overs'] * 6 / figures['wickets']
figures['bowling_average'] = figures['runs'] / figures['wickets']

#Displaying Bowler Figures

figures = figures.sort_values(['match_id', 'overs','runs'], ascending=[True, False, True])

figures['figure'] = (figures['overs'].astype(int).astype(str) + '-0-' + figures['runs'].astype(int).astype(str) + '-' + figures['wickets'].astype(int).astype(str))

bowlerset = set()
for i in figures['bowler']:
    bowlerset.add(i)

selected = figures[["match_id", "bowler", "figure", "economy", "fours", "sixes"]]

mergeddataset = selected.merge(df_match_info[['match_id', 'opponent']], on="match_id", how="left")

for bowler, bowler_df in mergeddataset.groupby("bowler"):
    print("\nBowler:", bowler)
    bowler_df = bowler_df.reset_index(drop=True)
    print(bowler_df[["opponent", "figure", "economy", "fours", "sixes"]])

#Phase Setup

df_bowling["phase"] = "Middle"
df_bowling.loc[df_bowling["over"] <= 6, "phase"] = "Powerplay"
df_bowling.loc[df_bowling["over"] >= 16, "phase"] = "Death"    

phase_figures = (
    df_bowling.groupby(['bowler', 'phase'], as_index=False).agg
    (
        overs = ('over', 'count'),
        runs = ('runs_conceded', 'sum'),
        wickets = ('wickets', 'sum'),
        fours = ('fours', 'sum'),
        sixes = ('sixes', 'sum'),
    )
        )

phase_figures["economy"] = (phase_figures["runs"] / phase_figures["overs"]).round(2)

#Displaying Phase-wise Bowler Figures

print("\nPhase-wise Bowler Figures:")

phase_order = ["Powerplay", "Middle", "Death"]
phase_figures["phase"] = pd.Categorical(phase_figures["phase"], categories=phase_order, ordered=True)
phase_figures = phase_figures.sort_values(["bowler", "phase"])


for bowler, bowler_df in phase_figures.groupby("bowler"):
    print("\nBowler:", bowler)
    bowler_df = bowler_df.reset_index(drop=True)
    print(bowler_df[["phase", "overs", "runs", "wickets", "economy", "fours", "sixes"]])


