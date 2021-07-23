# How to export multiple pages from OneNote to a single Word document – Cameron Dwyer

[https://camerondwyer.com/2017/04/13/how-to-export-multiple-pages-from-onenote-to-a-single-word-document/](https://camerondwyer.com/2017/04/13/how-to-export-multiple-pages-from-onenote-to-a-single-word-document/)

I needed to find a way to export a number of pages from a OneNote notebook into Word documents. The technique I used and will step through in this post was to:

- Create a new OneNote section (temporarily in needed) to arrange the pages I wanted to export (one section per Word document I wanted)
- Export the entire section into a Word (*.docx) file
- Automatically fix up image sizing issues with a custom macro

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb6.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb6.png)

Right-click on the section tab and select Export…

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb7.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb7.png)

Change the export file type to be a Word document (*.docx)

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb8.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb8.png)

This will export all pages from the OneNote section and append them all into one Word document. This is a pretty good result except for the pictures. In many cases the pictures are wider than the page width and look half missing.

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb9.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb9.png)

The solution to quickly and easily address the image width problem in bulk is to create a Word macro to resize all images in a Word document that are too wide to fit on the page.

Here’s how we create the Word macro (this is in Outlook 2016)

The Developer toolbar needed for creating macros isn’t visible by default so to switch it on to File | Options | Customize Ribbon and ensure Developer is checked

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb10.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb10.png)

You should now get the Developer toolbar appearing

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb11.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb11.png)

Select Macros, give the new macro a name and select a scope for where to save the macro (this determines where it will be available later on).

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb12.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb12.png)

You will now be dropped into the VB Macro Editor experience which looks nothing like Word! Don’t worry you just need to paste the following code in as shown

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb13.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb13.png)

Here’s the macro code to copy/paste

[Untitled](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/Untitled%20Database%20db95e7a06c204bba874de8a41692d9cd.csv)

The result should look like this, you can then save and close the macro editor window

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb14.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb14.png)

Now back in Word, with your document open (with the oversized images) select Developer | Macros, select your new macro and click Run

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb15.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb15.png)

Once the macro has completed all the images that were over 15cm in width will have been resized to fit on the page.

This assumes the pages are in portrait orientation and that the maximum width of an image should be 15cm

![How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb16.png](How%20to%20export%20multiple%20pages%20from%20OneNote%20to%20a%20sin%200bf479fe222441b781d630a8a66528a4/image_thumb16.png)