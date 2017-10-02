// Copyright (c) 2017 ActivePDF, Inc.
// ActivePDF DocConverter 2015
// Example generated 09/25/17 

using System;
using System.IO;
using System.Security.Cryptography;

// Make sure to add the ActivePDF product .NET DLL(s) to your application.
// .NET DLL(s) are typically found in the products 'bin' folder.

class Examples
{
    //Function to get random number
    public static int GetRandomNumber()
    {
        Random rnd = new Random();
        int n = rnd.Next();
        return n;
    }
    public static void DeleteFiles(FileInfo[] Files, string strPath)
    {
        //Delete the files
        for (int i = 0; i < Files.Length; i++)
        {
            if (File.Exists(@strPath + Files[i]))
            {
                File.Delete(@strPath + Files[i]);
            }
        }
    }  
   public static FileInfo[] FindFiles(string strPath, string filetype)
   {
        string str = "";
        
        //Find docx files 
        System.IO.DirectoryInfo d = new DirectoryInfo(@strPath);//Assuming Test is your Folder
        FileInfo[] Files2 = d.GetFiles("*" + filetype); //Getting Text files
        FileInfo[] Files = d.GetFiles("*" + filetype); //Getting Text files
        foreach (FileInfo file in Files)
        {
            str = str + ", " + file.Name;
        }
        return Files;
    }
  public static void Example()
  {
        string Wg = "WebGrabber \n 2016";
        string Dc = "DocConverter \n 2015";
        string Sv = "Server \n 2013";
        int PageCounter;
        int NumberOfPages;
        int intMergeFile;
        int intCopyForm;
        float textHeight;
        float textWidth;
        string strTitle = "First \n 25 \n pages "; ;
        int intOpenOutputFile;
        int intOpenInputFile;
        string strPath;
        string strPath2;
        
        DCDK.Results.DocConverterResult results;
        strPath = System.AppDomain.CurrentDomain.BaseDirectory;
        strPath2 = System.AppDomain.CurrentDomain.BaseDirectory + "title\\";

        FileInfo[] Files = FindFiles(strPath, ".docx");
        // Instantiate Object
        APToolkitNET.Toolkit oTK = new APToolkitNET.Toolkit();
        APDocConverter.DocConverter oDC = new APDocConverter.DocConverter();

        oDC.OverwriteMethod = ADK.Conversion.OverwriteMethod.AlterFilename;
   
        // Set the amount of time before a request will time out
        oDC.TimeoutSpan = new TimeSpan(0, 0, 40);
    
        // Enable extra logging (logging should only be used while troubleshooting)
        // C:\ProgramData\activePDF\Logs\
        oDC.Debug = true;
        //====================================================================================================================================
        //============================Word to PDF=============================================================================================
        // If the output parameter is not used the created PDF will use
        // the input string substituting the filename extension to 'pdf'
        for (int i = 0; i < Files.Length; i++)
        {          
            results = oDC.ConvertToPDF(strPath + Files[i], strPath + Files[i]+".pdf");
            if (results.DocConverterStatus != DCDK.Results.DocConverterStatus.Success)
            {
                ErrorHandler("ConvertToPDF", results, results.DocConverterStatus.ToString());
            }
         }
        //====================================================================================================================================
        //============================End Word to PDF=========================================================================================

        //Find the Newly Created Files for Stamping
        Files = FindFiles(strPath, ".pdf");

       //====================================================================================================================================
       //============================Stamping Logic==========================================================================================
       //For each PDF
        for (int i = 0; i<Files.Length; i++)
        {
            // Create the new PDF file
            intOpenOutputFile = oTK.OpenOutputFile(strPath + Files[i]+".pdf");
            if (intOpenOutputFile != 0)
            {
                ErrorHandler("OpenOutputFile", intOpenOutputFile);
            }
            //Get the number of pages of the pdf
            NumberOfPages = oTK.NumPages(strPath + Files[i]);
            NumberOfPages++;
            // Open the template PDF
            intOpenInputFile = oTK.OpenInputFile(strPath + Files[i]);
            if (intOpenInputFile != 0)
            {
                ErrorHandler("OpenInputFile", intOpenInputFile);
            }         
            textWidth = oTK.GetHeaderTextWidth(strTitle);
            textHeight = oTK.GetTextHeight(strTitle);
            PageCounter = NumberOfPages;
            PageCounter = 1;
            //For each page of the PDF do this
            for (int j = 1; j<NumberOfPages; ++j)
            {   
                //Determine which title stamp to use         
                if (i ==0)
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 110, 750, textWidth, textHeight, "Title: \n " + Dc, 1, j);
                    oTK.PrintMultilineText("Calibri |bold", 11, 500, 750, textWidth, textHeight, "Deposit \n Materials", 1, j);
                }
                else if(i ==1)
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 110, 750, textWidth, textHeight, "Title: \n " + Sv, 1, j);
                    oTK.PrintMultilineText("Calibri |bold", 11, 500, 750, textWidth, textHeight, "Deposit \n Materials", 1, j);
                  
                }
                else if(i ==2)
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 110, 750, textWidth, textHeight, "Title: \n " + Wg, 1, j);
                    oTK.PrintMultilineText("Calibri |bold", 11, 500, 750, textWidth, textHeight, "Deposit \n Materials", 1, j);
                }
                //First and Last 25 pages stamp
                if (j <26)
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 240, 32, textWidth, textHeight, "First \n 25 \n pages \n -", 0, j);
                }
                else
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 240, 32, textWidth, textHeight, "Last \n 25 \n pages \n -", 0, j);
                }
                //Page Count
                if(j < 26)
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 325, 32, textWidth, textHeight, "Page \n " + j, 0, j);
                }
                else
                {
                    oTK.PrintMultilineText("Calibri |bold", 11, 325, 32, textWidth, textHeight, "Page \n " + PageCounter, 0, j);
                    ++PageCounter;
                }           
            }
            // Copy the template (with the stamping changes) to the new file
            // Start page and end page, 0 = all pages
            intCopyForm = oTK.CopyForm(0, 0);
            if (intCopyForm != 1)
            {
                ErrorHandler("CopyForm", intCopyForm);
            }
            oTK.CloseOutputFile();
        }
        //====================================================================================================================================
        //============================End Stamping Logic======================================================================================      

        DeleteFiles(Files, strPath);

        //First Gather the newly created stamped PDFS
        Files = FindFiles(strPath, ".pdf");
               
        //Gather Files in title folder
        FileInfo[] Files2 = FindFiles(strPath2, ".pdf");
                  
        //====================================================================================================================================
        //============================Merge Title Page to stamped PDFs========================================================================
        //Merge the title Page
        for (int i = 0; i < Files.Length; i++)
        {
            // Create the new PDF file
            intOpenOutputFile = oTK.OpenOutputFile(strPath + Files[i] + "merged.pdf");
            if (intOpenOutputFile != 0)
            {
                ErrorHandler("OpenOutputFile", intOpenOutputFile);
            }
            // Set whether the fields should be read only in the output PDF
            // 0 leave fields as they are, 1 mark all fields as read-only
            // Fields set with SetFormFieldData will not be effected
            oTK.ReadOnlyOnMerge = 1;
            // Merge the cover page (0 for all pages)
            intMergeFile = oTK.MergeFile(strPath + "title\\"+Files2[i], 0, 0);
            if (intMergeFile != 1)
            {
                ErrorHandler("MergeFile", intMergeFile);
            }
            intMergeFile = oTK.MergeFile(strPath + Files[i], 0, 0);
            if (intMergeFile != 1)
            {
                ErrorHandler("MergeFile", intMergeFile);
            }
            // Close the new file to complete PDF creation
            oTK.CloseOutputFile();
        }
        // Release Object
        oTK.Dispose();
        //====================================================================================================================================
        //============================End Merge Title Page to stamped PDFs====================================================================

        DeleteFiles(Files, strPath);

        //Find the new merged PDFs and move them to output folder
        Files = FindFiles(strPath, ".pdf");

        //Create Random number for folder generation
        int captureRN = GetRandomNumber();

        for (int i=0; i<Files.Length; i++)
        {
            //If the file exists create new folder and put files in there
            if(File.Exists(strPath + "\\output\\" + Files[i]))
            {
                DirectoryInfo di = Directory.CreateDirectory(strPath + "\\output\\" + captureRN);
                File.Move(strPath + Files[i], strPath + "\\output\\" + captureRN +"\\"+  Files[i]);
            }
            else
            {
                File.Move(strPath + Files[i], strPath + "\\output\\" + Files[i]);
            }     
        }
        // Release Object
        oDC = null;  
    // Process Complete
    WriteResults("Done!");
  }
  
  // Error Handling
  public static void ErrorHandler(string strMethod, ADK.Results.Result results, string errorStatus)
  {
    WriteResults("Error with " + strMethod);
    WriteResults(errorStatus);
    WriteResults(results.Details);
    if (results.Origin.Function != strMethod)
    {
      WriteResults(results.Origin.Class + "." + results.Origin.Function);
    }
    if (results.ResultException != null)
    {
      // To view the stack trace on an exception uncomment the line below
      //WriteResults(results.ResultException.StackTrace);
    }
    Environment.Exit(1);
  }
  
  // Error Handling
  public static void ErrorHandler(string strMethod, object rtnCode)
  {
    WriteResults(strMethod + " error:  " + rtnCode.ToString());
  }
  
  
  // Write output data
  public static void WriteResults(string content)
  {
    // Choose where to write out results
  
    // Debug output
    //System.Diagnostics.Debug.WriteLine("ActivePDF: * " + content);
  
    // Console
    Console.WriteLine(content);
  
    // Log file
    //using (System.IO.TextWriter writer = new System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory + "application.log", true))
    //{
    //    writer.WriteLine("[" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + "]: => " + content);
    //}
  }
}