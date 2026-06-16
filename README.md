# Trivia Night — slides

Slide deck for the **16 June trivia night** at the Permanent Representation of
Lithuania to the EU, built from the questions drafted in
`Klausimai protmūšiui (2).docx`.

Format: **one slide — question, next slide — answer.** The deck is fully in
English for an international audience. Production notes from the draft
(source links, hint options, alternative question versions) live in the
**speaker notes** of each slide.

## Files

| File | Purpose |
|---|---|
| `Trivia_Night_2026-06-16.pptx` | The deck (60 slides, 16:9) — open in PowerPoint |
| `Trivia_Night_2026-06-16.pdf` | Static preview — **re-export from PowerPoint** after rebuilding; a PDF cannot play the embedded audio/video |
| `Trivia_Night_2026-06-16_run_of_show.docx` | Run of show: every round/question/answer in presentation order, plus the changes made relative to the draft |
| `Trivia_Night_2026-06-16_answer_key.docx` | Printable A4 answer key (brief question + answer + points per round) for marking team answers |
| `team_cards/table_signs.docx` | A4-landscape table tents — one team name per page in the bottom half only, with a dashed fold line; fold in half so the name shows on one face (Lithuanian rivers/lakes) |
| `team_cards/draw_slips.docx` | A4 draw slips — small, uniform slips, five per team (one per member), all 50 on a single page to cut up and draw at the door |
| `generate_slides.py` | Builds the deck: `pip install python-pptx Pillow && python3 generate_slides.py` |
| `generate_rundown.py` | Builds the run-of-show docx: `pip install python-docx && python3 generate_rundown.py` |
| `generate_answer_sheet.py` | Builds the answer key: `pip install python-docx && python3 generate_answer_sheet.py` |
| `generate_team_cards.py` | Builds the team signs and draw slips: `pip install python-docx && python3 generate_team_cards.py` |
| `fetch_assets.py` | Downloads/renders the illustrations into `assets/`: `pip install cairosvg && python3 fetch_assets.py` |
| `assets/` | Rendered illustrations (committed, so the deck builds offline) |
| `photos/` | Photographs and album art (`song_*`) used on the slides |
| `media/audio/` | Song excerpts embedded in the Music round |
| `media/video/` | Video clips embedded in the deck |
| `Klausimai protmūšiui (2).docx` | Latest question draft (source for the deck) |

## Design

- Bold dark theme: near-black slides, with one saturated colour per round —
  EU blue for Round 1, then the Lithuanian tricolour (yellow, green, red)
  for Rounds 2–3 and the bonus round.
- Full-bleed colour divider slides with oversized numerals; question slides
  carry a giant question number and a colour-coded edge bar; answers are
  revealed in the round colour.
- Sharp, rectangular geometry throughout — no rounded corners.
- **Segoe UI** typography (Segoe UI Black for the display headings) — a
  font installed by default on Windows, so nothing has to be embedded and
  managed/locked-down Office installs render the deck as intended.
- Real photographs (`photos/`, see `photos/CREDITS.md`) illustrate most
  questions and answers, placed at native aspect ratio — never cropped —
  and sized to fill the right side of the slide; remaining icons are
  [OpenMoji](https://openmoji.org) (CC BY-SA 4.0).

## Deck structure

1. **Title + rules** — team draw at the door (slips with Lithuanian rivers/lakes)
2. **Round 1 — Europe & Fun Facts** · 8 multiple-choice questions
3. **Round 2 — Lithuania** · 8 questions
4. **Round 3 — Environment & Nature** · 6 questions
5. **Round 4 — Music (Name the Connection)** · 5 riddles whose answers are
   all about the sea — a nod to the EU's upcoming Ocean Pact (and to
   Čiurlionis's *Jūra*); the matching song excerpt is available
   from the embedded player on each slide as the clue
6. **Bonus — The Brussels Bubble** · 3 insider questions (1–2 points)
7. **Tie-breaker** · one question (Napoleon / St Anne's Church), used only
   if teams finish level
8. **Closing / scoring**

## Embedded audio & video

All clips are **embedded in the .pptx and play inline** in PowerPoint
(offline). Each sits behind an on-brand poster frame; click the play
button to start. **Videos are set to play full screen.** A **hyperlinked
title sits under each video** (and the URL is in the speaker notes) so you
can fall back to YouTube if a clip won't play on the night.

- **Music round** — five excerpts in `media/audio/` (Yellow Submarine,
  Ocean Eyes, How Far I'll Go, My Heart Will Go On, Jūra). Each riddle has a
  small embedded player in the corner; click it to hear the clue, and the
  song and its album art are revealed on the answer slide.
- Round 2 Q1 **answer** — LT United, *We Are The Winners*
  (`media/video/eurovision_we_are_the_winners.mp4`). The clip lives on the
  **answer** slide so it doesn't give the title away; the **question** keeps
  just the Eurovision logo. Trim to the part you want in PowerPoint
  (Playback ▸ Trim Video).
- Round 2 Q7 — EU-accession promo. The **question** slide plays a short
  clip cut to 26 seconds
  (`media/video/eu_accession_spirgi_clip.mp4`) — it stops at the line about
  membership opening new markets, before the EU is named, so it doesn't give
  the answer away. The **answer** slide reveals the full promo
  (`media/video/eu_accession_spirgi.mp4`). Add English subtitles before the
  night.
- Round 1 Q5 answer — the Umaru Dikko story
  (`media/video/umaru_dikko_kidnap.mp4`).

> The clips are sized to a 16:9 frame. If a source clip has a different
> aspect ratio it will letterbox/stretch — drag the frame's handles in
> PowerPoint to adjust.

## Question pool not yet on slides (from the draft — to discuss)

- CBAM question about nails (needs wording)
- The Schumann Show / the institutions ([video](https://www.youtube.com/watch?v=D0fBh-0Eiy0))
- One question per member state / quick-fire round idea
- Questions voiced by the Minister and Selemonas Paltanavičius (incl. the Belém fire)
- Vytis; the Pacai–Pazzi family link; Bona Sforza
- German street in Vilnius
- Kaunas modernism / Vilnius baroque
- Adamkus 100 / Attenborough 100; the bears offered by Slovenia
- Baltic Way symbols in the European Parliament
