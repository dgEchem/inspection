# Steel Member Survey (iPad app)

Application for surveying steel members (I-beams,
L brackets, C brackets) with photos, ultrasonic thickness readings and notes.
Same hosting, install and backup story as the modillion app - see its README
for GitHub Pages and "Add to Home Screen". It is a separate app with its own
address and its own data; install both on an iPad if you want both.

Files:

    index.html / sw.js / manifest.webmanifest / icon-*.png   the app
    member_report.py                                         PC tool: merge exports, build the Word report

## Capture tab

* **Area** and **Part #**. Next keeps the
  area and the part type and increments the number.
* **Part type** dropdown (I-beam / L bracket / C bracket / Other). The lettered
  diagram and the surface list come from the type.
* **Overall photo** - the wide shot. The `+` adds more overall photos; extras
  stay as small thumbnails here and go in the report as a row of small photos
  under the main one.
* **Surfaces being surveyed** - tap the letters you are measuring (A, C, ...).
  A section opens below for each one with:
  * a **Surface photo** tile,
  * a **Gauge screen** tile (optional - a photo of the 39DL PLUS display),
  * **Thickness readings** - type or paste one or several values
    (`0.512, 0.498 0.505`); min and average are shown and go in the report,
  * **quick-pick defects** and **notes** for that surface.
* **General notes** for the part at the bottom.

Everything saves instantly (Report tab > Export project zip; big projects export in
numbered files).

## Report

A page per part: grey header bar
"Area 1: Part 3 (I-beam)", the overall photo with the general notes beside it,
a row of small extra overall photos if any, then one row per surveyed surface
with the surface photo on the left and, on the right, the thickness readings
(min / avg), the notes, and the gauge photo if one was taken. Captions carry
the surface letter and description ("Figure 5: Part A1-1 - Surface A (Top
flange, top)"). A "Summary of defects by area" table (Area / Part / Surface /
Defects) opens the document. Tables flow over the page when a part has many
surfaces.

## Settings

Project name, operator, dark mode, footer text, element/group names ("Part" /
"Area"), thickness units (in / mm), section number, first figure number,
photo sizes, export part size, quick-pick defects, and the **part types**:

    Type name | code | A:description; B:description; ...

Edit the descriptions to match how the crew refers to the surfaces; the
letters are what the checkboxes show. The diagrams are fixed for codes I, L
and C.

## PC: merging exports and building the report

    pip install python-docx pillow
    python member_report.py --merge Merged export1.zip export2.zip ...
    python member_report.py --build Merged/project.json Report.docx
    (or both: python member_report.py --merge Merged a.zip b.zip --build Report.docx)

Zips are unpacked into an `_unpacked` folder next to them. Duplicates (same
area + part number across exports) are resolved automatically: the first copy
is kept unless a later one is more complete (more photos / readings / notes).

## Ultrasonic data (39DL PLUS)

For now readings are typed or pasted, and the gauge display can be
photographed. The readings field is plain text, so when a CSV export from the
gauge (Link Plus app / GageView) is available it can be pasted in or - with a
small addition to the app - imported and matched to parts by the ID numbers
entered on the gauge (use IDs like `A1-3-A` = area 1, part 3, surface A).
