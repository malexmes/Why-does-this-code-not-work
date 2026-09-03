# The Econometrics Game, version 2

A no-macro Excel workbook that teaches non-technical marketing teams how the Analytic Partners Wave 1 model of Race for Life sign-ups works, one block at a time.

## What is in the workbook

| Tab | What it does |
|---|---|
| Start here | The detective picture, how to play, two routes (five minutes or thirty), progress tracker |
| Level 1 to Level 10 | One idea per tab. Regulars, brakes, the sale, the kettle (adstock), the halo, the channel split, the four tests, the next pound, brand and the long game, Committed Giving |
| Words | Every term on a three-rung ladder: everyday picture, what it meant in Wave 1, the sentence to say in the room. Four native charts |
| Under the bonnet | Every weekly number and formula the game uses. Visible on purpose |
| Room notes (hidden) | Timings, the answer key and the six questions a room will ask |
| Answers (very hidden) | The answer key the checks compare against |

## Where the numbers come from

Every weekly series is read from the charts in the Wave 1 results deck: real sign-ups, spend by channel and product, AP's weekly decomposition of sign-ups by driver, brand consideration and new Committed Givers. The answer key is fitted to that decomposition. Nothing is illustrative except the four small pictures on the Words tab, which are labelled as such.

## Rebuilding for the next wave

```
pip install python-pptx openpyxl numpy
python build_econometrics_game.py <Wave_results.pptx> CRUK_Econometrics_Game_v2.xlsx
```

Then recalculate once in Excel or LibreOffice. The script prints the fitted answer key and the pattern-explained targets for each level.
