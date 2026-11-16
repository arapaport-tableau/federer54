# Federer 54

> "In the 1,526 singles matches I played in my career, I won almost 80% of those matches. Now, I have a question for all of you: What percentage of points do you think I won in those matches? Only 54%." — Roger Federer, Dartmouth commencement, 2024

A data story about Roger Federer's career: how a 54% edge in points became one of the most dominant records in sport. It covers his career arc, the margin, leverage, surfaces, shot-by-shot analysis, how his game evolved, pressure, and replays of his defining matches.

Federer 54 was built as a **Tableau Data App**. This repository holds a standalone web version, so it runs on GitHub Pages without Tableau.

## Run it

Open `index.html` in a browser, or serve the folder with any static web server. No build step.

## Files

- `index.html`: the page
- `app.js`, `styles.css`: the app
- `data.js`: the derived data, plus a small stand-in for the Tableau Extensions API
- `photos.js`: photo list (empty in the public version; see `photos.example.js`)

## Data and credits

- Match Charting Project and Grand Slam point-by-point data: Jeff Sackmann / Tennis Abstract and contributors, licensed CC BY-NC-SA 4.0. Derived data in this repository is shared under the same licence.
- Career match results and statistics: TennisMyLife (personal, non-commercial use).
- Speeds are radar readings of serves from Grand Slam data (2011 onward). Rally-shot speeds are estimates and are labelled as such.

For non-commercial, educational use only. Not affiliated with Roger Federer, the ATP or any tournament.
