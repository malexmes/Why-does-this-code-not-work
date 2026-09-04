# The Econometrics Game

Excel games that teach non-technical marketing teams how the Analytic Partners Wave 1 model of Race for Life sign-ups works.

## Version 3 (current): one screen, like an app

`CRUK_Econometrics_Game_v3.xlsm` is a single-screen game. A progress bar, one picture, one sentence, one control, one Next button. Twenty-one steps in seven chapters on the Play screen, then a budget game, two short reading chapters and a finish screen with ten sentences to say in the room. Ten quiz stars.

Buttons are cells. Clicking one fires a small VBA handler that moves the game state on the hidden State tab. All the maths stays in formulas, so `CRUK_Econometrics_Game_v3_no_macros.xlsx` is the same game driven from a visible State tab for machines that block macros.

| Tab | What it does |
|---|---|
| Play | Chapters 1 to 7: receipts, the shop, the sale, the kettle, the next field, who sang louder, any good |
| Budget | Chapter 8: move money between the thirteen 2025 channels, keep the total, beat the plan |
| Long game | Chapter 9: brand consideration and the long-term multiplier |
| Giving | Chapter 10: why Committed Giving shows 9p per £1 |
| Finish | Stars and the ten sentences |
| Words | Every term on a three-rung ladder |
| Under the bonnet | Every weekly number and formula. Visible on purpose |
| Steps, State (hidden) | The step content table and the game state |
| Room notes (hidden) | Timings, answer key, the questions a room will ask |
| Answers (very hidden) | The answer key the checks compare against |

## Where the numbers come from

Every weekly series is read from the charts in the Wave 1 results deck: real sign-ups, spend by channel and product, AP's weekly decomposition by driver, brand consideration and new Committed Givers. The answer key is fitted to that decomposition.

## Rebuilding

```
pip install python-pptx openpyxl numpy oletools
python build_econometrics_game_v3.py <Wave_results.pptx> app_v3     # writes app_v3.xlsx
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" &   # LibreOffice with python3-uno
python add_vba.py app_v3.xlsx CRUK_Econometrics_Game_v3.xlsm          # exports the VBA project through LibreOffice and injects it
python test_vba.py                                                     # drives every button through the VBA in LibreOffice
```

`Game.bas` is the VBA. `CRUK_Econometrics_Game_v2.xlsx` is the previous tab-per-level version, kept for reference.
