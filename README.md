# gif2xlsx
Convert GIFs to XLSX format

I worked for a long time in financial services, and I can tell you that one thing I got sick of hearing was "how do I convert animated GIFs into Excel files". If you're here wondering why nobody has yet exploited this gap in a very lucrative market, wonder no longer. Help is at hand.

![BluePrint build order](https://github.com/pugwonk/gif2xlsx/raw/master/readmepics/sample.png)

# Usage

1. Download [gif2xlsx.exe](https://github.com/pugwonk/gif2xlsx/releases)
1. Download your favourite GIF to the same folder
1. At a command line in the same folder, type: `gif2xlsx myfavourite.gif` (ot whatever you called it)

Animated GIFs are converted on a one-frame-per-worksheet basis, so you have to step through the worksheets to animate them. I was originally intending using conditional formatting and iterative calc to display these, but unfortunately the Excel team* seem to have single-threaded the calculation of conditional formatting and it was far too slow to render.

\* I used to work on the Excel team so one could argue that this is partly my fault.

# FAQ
* How does this work?
	* It uses Microsoft's [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) to generate Excel files and .NET's [System.Drawing.Image](https://msdn.microsoft.com/en-us/library/system.drawing.image(v=vs.110).aspx) to create the workbook. As GIF is a colour-indexing format, the program maintains a palette to avoid generating loads of extra formatting records. Because nobody wants their animated spreadsheets to be too large to email to colleagues.
* Why is it so slow?
	* The [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) isn't all that fast when you write cell values. There are [much faster SAX-style ways of writing out spreadsheets using the same SDK](https://blogs.msdn.microsoft.com/brian_jones/2010/05/27/parsing-and-reading-large-excel-files-with-the-open-xml-sdk/) but it's harder and I couldn't be bothered. Also, let's be honest, if you've got time to convert GIFs to spreadsheets then you've sure got time to wait for it to finish.
* Why doesn't my GIF scale properly?
	* The program scales all GIFs to a fixed size. This could probably be fairly easily fixed but I was really supposed to be working today instead of doing any of this.
* Why does it crash out when given an invalid GIF filename, or when exposed to direct sunlight?
	* Because error handling is boring, and when my wife says "what have you been up to today" there needs to be more for me to say than "well, I wrote this converter".
* Why didn't you make this some sort of web service?
	* It's on GitHub - be my guest! You are welcome to 100% of the profits.
* Why is your release binary built in debug mode?
	* Be grateful for what you have.
	
