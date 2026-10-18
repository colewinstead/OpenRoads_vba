# Pattern Line toolbar icon

A cyan guide line with four orange perpendicular pattern lines.

## Recommended files after an invisible custom icon

Use **PatternLineGenerator-Bentley.ico**. It contains uncompressed 24-bit Windows
DIB images at 16, 24, 32 and 48 pixels, with an opaque dark gray background.
It avoids PNG-compressed ICO images and alpha transparency. The original
**PatternLineGenerator.ico** used PNG-compressed transparent frames; this is a
possible import compatibility issue, not a confirmed diagnosis in ORD. The user
reported both original ICO and BMP imports becoming invisible after Apply, so
PNG encoding alone does not explain the reports. Verify icon storage and the
saved button reference in the customization library.

### Import the icon into the customization library

First note the library named in the **Customize Ribbon** title, then open that
actual **.dgnlib** in ORD. For personal ribbon buttons this is usually
**Personal.dgnlib**. Importing while a normal drawing is open may store an icon
in a different file from the ribbon customization. Bentley explicitly opens
the personal library before importing icons in its
[personal custom button instructions](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0100334).

1. With the customization **.dgnlib** open, choose **Customize Ribbon**, then
   **Custom Tools and Tasks** at the bottom left.
2. On the **Icons** tab, open the plus dropdown and choose **From Ico Files**.
3. Import **PatternLineGenerator-Bentley.ico** using a new icon name such as
   **PatternLineGenerator_Bentley** to distinguish it from the previous import.
4. Return to the ribbon button's **Icon** property. In **Define Icon for Tool**,
   set **Look for icons in** to **Current Design File**, then select the newly
   imported icon. Click **OK**, then **Apply**.

Use the customization library already containing the ribbon button.
This follows Bentley's [CONNECT ribbon customization instructions](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012777).
No workspace variables or VBA code need changes for this test.

### Bitmap fallback

For **Import Selected Bitmaps**, browse to **PatternLineGenerator-Bentley-32.bmp**.
It is an uncompressed 32 x 32, 24-bit bitmap with a one-pixel white border and a
dark gray tile. Bentley documents using the upper-left pixel as the transparency
key during BMP import, with a white border to control that key:
[bitmap transparency guidance](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0039731).

## Original artwork

**PatternLineGenerator.png** is the transparent full-size artwork;
**PatternLineGenerator-32.png** is the transparent 32 pixel preview.
The original ICO and BMP are retained for comparison.

Artwork created with the built-in image generation tool. Design prompt:
A simple flat CAD toolbar pictogram with one straight diagonal cyan-blue baseline,
four evenly spaced orange-gold crossbars at 90 degrees, thick square-ended strokes,
a subtle dark outline, transparent background, and no text or extra symbols.
The compatibility versions use the same artwork with a dark gray background.

All four ICO frame headers, image sizes, uncompressed 24-bit format, opaque masks,
and bitmap header/border were checked locally. The preview was visually checked.
Import and persistence after Apply in ORD remain untested.
