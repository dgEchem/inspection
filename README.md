# Steel Member Survey - iPad app (v1.3.0)

A field app for steel member surveys (I-beams, L and C brackets, anything marked on a
photo). For each part you take an overall photo, tick the surfaces you measured, take a
photo of each surface (and of the gauge screen), type the thickness readings, drop
numbered points on the photo where each reading was taken, and note defects. At the end
the app builds the Word report on the iPad, or you export a zip and build it on the PC
with `inspection_report_app.py`.

Everything is stored on the iPad itself (no account, no server). It works with no signal
once it has been opened once online.

## Installing / updating

1. Put `index.html`, `sw.js`, `manifest.webmanifest` and the three `icon-*.png` files in a
   GitHub Pages folder (the `README.md` and `member_report.py` are not needed there).
2. On the iPad open the page in Safari -> Share -> **Add to Home Screen**. Always launch it
   from the home-screen icon (a normal Safari tab can be cleared by iOS; the icon is the
   safest place for the data).
3. To update: push the new files, then open the app **while online** and close/reopen it
   once or twice. The version number shows at the bottom of Settings. Projects, photos and
   undo history stay on the iPad through an update.

## Capture screen

* **Area / Part # / Part type / Units** - Next carries the area, type and units forward and
  adds 1 to the number. Units (in / mm / mils) are saved per part.
* **Overall photo** - one main wide shot plus small extra overall photos.
* **Surfaces** - tap the diagram (red = selected) or the lettered buttons; a section opens
  for each selected surface with its photo, gauge-screen photo, measurements, defect
  quick-picks and notes. "Marked on photo" parts have no fixed surfaces: tap the overall
  photo to drop A, B, C... and describe each one.
* **Measurements** - Add measurement gives a numbered box (value + optional note). Each row
  has a **Place / Move** button: it asks first, then you tap the surface photo to put (or
  move) that number there. Numbers on the photo are orange when placed, grey when not.
  "Place new on photo" adds a new reading at the spot you tap.
* **Photos are checked** - after a photo is saved the app reads it back; if that fails you
  get a red warning while the person is still standing at the part, so retake it then.

## Undo / redo (v1.3.0)

The header has **Undo**, **Redo** and a **clock** button (recent changes).

* Everything that changes a part can be undone: un-selecting a surface (with its photo and
  measurements), deleting a measurement, deleting or replacing a photo, deleting a part,
  moving a point, editing text, importing a zip, changing part type or units. Area
  introductions are covered too.
* Risky actions (removing a selected surface that has data, deleting a part, deleting a
  photo) show a message with an **Undo** button for about 7 seconds. Tapping Undo restores
  the surface with its photos and every measurement.
* Typing is grouped: a burst of typing in one box is one undo step, not one per letter.
* The **clock** lists the recent changes; tap **Undo to here** to roll back several steps.
* Photos are not truly deleted straight away. A removed photo stays on the iPad until nothing
  can restore it any more (the undo history is kept when you close and reopen the app).
  Only then is it deleted and the space reclaimed. Anything that has already been exported
  is of course in the zip.
* Keyboard: Ctrl/Cmd+Z undoes and Shift+Ctrl/Cmd+Z redoes (not while typing in a box - the
  keyboard's own undo works there).
* Limits: the last 600 steps are kept on this iPad; "Start a new project" clears the history;
  app settings (names, photo size, defect list) are not part of undo.

## Report tab

* **Generate Word document** - whole report or one area. Keep the app on screen while it
  builds. Then **Share / Save to Files** (choose *On My iPad* or *OneDrive*) or Download.
* **Export project zip** - all photos plus `project.json`. Large projects are split into
  numbered parts (default 300 photos each); export every part. The line under the buttons
  shows how much storage is used, how much is left and about how many more parts fit.
* **Import zip (merge)** - merges another iPad's export. Parts are matched by area + part
  number; an existing part is kept unless the incoming copy is more complete. The import is
  one undo step. Photos the zip does not contain are reported, not ignored.
* **Start a new project** - erases everything on this iPad (it tells you how much has not
  been exported yet and makes you type ERASE).

### Back up as you go

Export a zip about once an hour and move it off the iPad (Files -> OneDrive, mail, AirDrop).
If something happens to the iPad or the browser data, the zip rebuilds the whole project.

## Other tabs

* **All parts** - the list of every part; tap one to open it. A duplicate area + part number is flagged on the capture screen.
* **Areas** - the heading page for each area: an intro sentence and an optional plan photo.
* **Settings** - project name, your initials, dark mode, report footer, section and figure
  numbers, element/group names, default units, photo size and quality, the part-type list
  (`Name | code | A:description; B:description; ...`) and the defect quick-picks
  (`Button label | sentence added to the notes`).

## On the PC

`inspection_report_app.py` (separate download) opens the exported zips, merges several
iPads, lets you fix parts and add photos from a camera, and builds the same Word layout.
`member_report.py` is the older command-line version of that and still works.

    python inspection_report_app.py export.zip
    python inspection_report_app.py --merge Merged a.zip b.zip --build Report.docx

## If something looks stale

If the app still shows an old version number after an update, make sure the new files are
live (open `.../sw.js` in the browser and check the `sms-x.y.z` name), then open the app
online and close it fully (swipe it away) twice. Your data is not affected.
