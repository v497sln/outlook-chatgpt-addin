# SquadSage: an offline FUT companion

SquadSage is a rebuild of the *idea* behind FUTGenie (an SBC solver, squad builder and transfer-list tracker for EA FC Ultimate Team). It was redesigned around what users complain about in that app's reviews. It is an independent fan tool and is not affiliated with EA or FUTGenie.

## What users said → what changed

| Review finding | Source | SquadSage's answer |
|---|---|---|
| 👍 "Saves loads of time on SBCs" (a few SBCs in under 5 minutes instead of 30) | Store reviews | Fast local solver (1.5 s, or a 4 s "thorough" mode) with ready-made challenge templates and your own saved SBCs |
| 👍 Auto solver, Gauntlet/squad builder, smart relist, profit tracking | Store listing, release notes | Solver, squad builder with live chemistry, and a trade tracker with "Relist all" and adjustments |
| 👎 A failed SBC submission still uses up one of the day's credits | Google Play review | **No credits at all.** Solving is unlimited, and a squad is shown as valid only after every requirement is checked |
| 👎 It stopped submitting SBCs, or features were down for hours even on the paid plan | Store reviews | No server is involved. Everything runs on the device, and the app works **offline** as an installable web app |
| 👎 "Will it hack my FC account?" Ban worries (one user reports a ban) | App Store, X replies | **No EA login and no automation.** You apply the squad yourself in-game. Nothing reaches your account |
| 👎 Premium "is just another way to get you to buy it". The price is 9.99 in the app but 5.49 on the website | Store reviews | Free, with no ads and no tiers |
| 👎 Confusing transfer dashboard (labels like "Realised income", unclear about tax) | Original app screenshot | Every figure is labelled in plain words and **includes the 5% EA tax**. Each trade shows its break-even price and you get a warning before listing at a loss |
| Missing: a way to stop the solver using players you want to keep | Common request | 🔒 **Protect** any card and it is never used |
| Missing: telling your club apart from what you'd have to buy | | The result splits "From your club" from "Buy from market" and shows the coins you'll spend. Untradeables and duplicates are used first |
| Rating rounding causes false "valid" squads | | Team rating is always rounded **down**, so the app never says pass when the game would reject it |

## Features

- **SBC solver**: requirements for team rating, squad chemistry, per-player chemistry, "min/max/exactly N players with nation/league/club/rarity/quality", "from the same …", "different …" and player rating range. Solving uses simulated annealing in a Web Worker. Cost rules: untradeables are free, duplicates are discounted, tradeable cards are valued at their after-tax price, and market fodder is costed from an editable price table. "Try another solution" gives you alternatives. "I submitted it" removes the used cards from your club and can be undone.
- **Squad builder**: 6 formations, live team rating and the 0–3 chemistry per player (club 2/5/8, nation 2/5/8, league 3/5/8; icons and heroes handled).
- **Club**: CSV import/export, manual add/edit, search and filters, protect toggle, duplicate counts, untradeable flag.
- **Trades**: track buys and listings, mark items sold or expired, relist all with ±% that never drops below break-even, and profit after tax.
- **Backup**: export and restore everything as JSON. Data is stored only in the browser.

## Run

```bash
cd fut-companion
npm start          # serves on http://localhost:8080 (any static server works)
npm test           # engine unit tests (node --test)
```

Open it on your phone and use "Add to Home screen" to install it.

`node build.mjs` bundles everything into one file, `dist/squadsage.html`, which you can open or host anywhere.

### CSV format

```
name,rating,positions,nation,league,club,rarity,untradeable,count,price,protected
Jane Doe,84,CB|RB,England,Premier League,Some FC,rare,yes,1,0,no
```

Only `name` and `rating` are required. Rows that can't be read are reported, not silently dropped.

## Limitations (honest)

- There are no live market prices. The app does not scrape prices or log into anything, so keep the fodder price table (⚙︎ Settings) up to date yourself.
- Importing your club is manual (CSV). This is the trade-off for never touching your account.
- The chemistry and rating formulas follow the publicly documented rules. Special cases in a new game edition may differ.
- The sample club is fictional.
