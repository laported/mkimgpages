var FSO = new ActiveXObject("Scripting.FileSystemObject");
var response = WScript.StdOut
var inputDir = WScript.Arguments.length > 0 ? WScript.Arguments(0) : ".";

makeHtmls(FSO.GetFolder(inputDir));

function makeHtmls(folder)
{
	var htmlFile;
	var sorted = new Array(1000);
	var files = new Enumerator(folder.Files);
	var numfiles;
    
	response.WriteLine("Input dir =" + inputDir);
    
	response.WriteLine("Sorting files...");
	for(i=0;!files.atEnd();files.moveNext())
	{
        if (files.item().Name.search(".jpg") != -1)
		{
			sorted[i] =  files.item().Name;
			response.WriteLine(files.item().Name);
			i++;
		}
	}
	sorted.sort();
	numfiles = i;
	response.WriteLine(i + " files sorted");
    
	response.WriteLine("Creating html files...");
	for(i=0;i<numfiles;i++)
	{
		lastfile =  sorted[i];
	}
    
	for(i=0;i<numfiles;i++)
	{
		if (i == 0)
		{
			htmlFile = folder.CreateTextFile("index.html", true) 
			htmlFile.WriteLine("<html><body><a href='" + sorted[i] + ".html'><img src='" + sorted[numfiles-1] + "'/></a></body></html>");
		}
		else if (i == numfiles-1)
		{
			htmlFile = folder.CreateTextFile(sorted[i-1] + ".html", true) 
			htmlFile.WriteLine("<html><body><a href='index.html'><img src='" + sorted[i-1] + "'/></a></body></html>");
		}
		else
		{
			htmlFile = folder.CreateTextFile(sorted[i-1] + ".html", true) 
			htmlFile.WriteLine("<html><body><a href='" + sorted[i] + ".html'><img src='" + sorted[i-1] + "'/></a></body></html>");
		}		
		htmlFile.Close();
	}
    
	response.WriteLine("OK.");
}

