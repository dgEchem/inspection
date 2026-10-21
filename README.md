# Steel Member Survey (iPad app) - v1.2.0

Sister app of the Modillion Field Report, for surveying steel members (I-beams,
L brackets, C brackets, or anything irregular) with photos, ultrasonic
thickness measurements and notes. Same hosting, install and backup story as
the modillion app - see its README for GitHub Pages and "Add to Home Screen".
It is a separate app with its own address and its own data; install both on an
iPad if you want both.

Files:

    index.html / sw.js / manifest.webmanifest / icon-*.png   the app
    member_report.py                                         PC tool: merge exports, build the Word report

## Capture tab

* **Area** and **Part #** replace zone and modillion number. Next keeps the
  area, the part type and the units, and increments the number.
* **Units** (in / mm / mils) at the top of each part. Whatever you pick stays
  for the following parts until you change it. Each part remembers its own
  units, so a project can mix them.
* **Part type** dropdown (I-beam / L bracket / C bracket / Marked on photo /
  Other).
  * For I-beam, L and C bracket the **diagram is clickable**: tap a face on the
    drawing (it turns red) or tap the lettered button - both do the same
    thing. A section opens below for each selected surface.
  * **Marked on photo** is for irregular things (the Excalibur monument): take
    the overall photo, tap *Add point*, then tap the photo where each surveyed
    spot is - A, B, C... appear on the photo and a section opens below for
    each, with a short "where is it" description for the caption. Tap a letter
    on the photo to remove it. The report shows the overall photo full width
    with the letters drawn on.
* **Overall photo** - the wide shot. The `+` adds more overall photos; extras
  stay as small thumbnails here and go in the report as a row of small photos
  under the main one.
* Each surface section has:
  * a **Surface photo** tile and a **Gauge screen** tile (optional - a photo
    of the 39DL PLUS display),
  * **Thickness measurements** - tap **+ Add measurement** for each reading;
    each is a numbered box with the value and an optional short note ("near
    weld"). Tap **Place on photo** and then tap the surface photo to put the
    numbered marker where that reading was taken (orange numbers; tap a
    marker to remove it). Min and average are shown and go in the report.
  * **quick-pick defects** and **notes** for that surface.
* **General notes** for the part at the bottom.

Everything saves instantly; the backup line and export rules are the same as
the modillion app (Report tab > Export project zip; big projects export in
numbered files). Projects saved by the previous version open fine - typed
readings like "0.512, 0.498" become numbered measurements automatically.

## Report

A page per part, same frame as the modillion report: grey header bar
"Area 1: Part 3 (I-beam)", the overall photo with the general notes beside it
(full width with letters for marked parts), a row of small extra overall
photos if any, then one row per surveyed surface with the surface photo on
the left (numbered orange points drawn on if you placed them) and, on the
right, "Thickness (in):" followed by the numbered list of measurements with
their notes, the min / average, the surface notes, and the gauge photo if one
was taken. Captions carry the surface letter and description ("Figure 5: Part
A1-1 - Surface A (Top flange, top)"). A "Summary of defects by area" table
(Area / Part / Surface / Defects) opens the document. Tables flow over the
page when a part has many surfaces.

## Settings

Project name, operator, dark mode, footer text, element/group names ("Part" /
"Area"), default units for new projects, section number, first figure
number, photo sizes, export part size, quick-pick defects, and the **part
types**:

    Type name | code | A:description; B:description; ...

Edit the descriptions to match how the crew refers to the surfaces; the
letters are what the buttons show. The diagrams are fixed for codes I, L and
C; code M is "marked on photo"; any other code gets buttons only.

## PC: merging exports and building the report

    pip install python-docx pillow
    python member_report.py --merge Merged export1.zip export2.zip ...
    python member_report.py --build Merged/project.json Report.docx
    (or both: python member_report.py --merge Merged a.zip b.zip --build Report.docx)

Zips are unpacked into an `_unpacked` folder next to them. Duplicates (same
area + part number across exports) are resolved automatically: the first copy
is kept unless a later one is more complete (more photos / measurements /
notes). The PC report is identical in layout to the one the iPad builds.

## Ultrasonic data (39DL PLUS)

For now measurements are typed and the gauge display can be photographed.
When a CSV export from the gauge (Link Plus app / GageView) is available it
can - with a small addition to the app - be imported and matched to parts by
the ID numbers entered on the gauge (use IDs like `A1-3-A` = area 1, part 3,
surface A; the measurement numbers 1, 2, 3 map to the gauge's sequence within
that ID).
