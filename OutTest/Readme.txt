###################################################################################################################################
1> Install Default Instance of Printer

    # Open Config.xml file.
	
	# Change PrinterName, PortName, HardwareId, outputFilePath,outputFileName,paperSize,color parameters Values as needed.
	
		* Here PrinterName,PortName,HardwareId must be different for all instances of printer.
			i.e. For Default Instance(First Instance) Below Values are Okay.
		
	                    <Parameter Name="PrinterName" Value="PdfScribe" />
                            <Parameter Name="PortName" Value="PSCRIBE" />
                            <Parameter Name="HardwareId" Value="PDFScribe_Driver0101" />

			            <Parameter Name="outputFilePath" Value="D:\\PDF_SCRIBE\\Output" />			
			            <Parameter Name="outputFileName" Value="PdfDoc.pdf" />
			            <Parameter Name="paperSize" Value="A0" />
			            <Parameter Name="color" Value="Color" />

	# Save Config.xml File	
	
	# Run PdfScribeInstall.msi Directly to install default instance of printer. 
	
	# This Will Install Scribe.


###################################################################################################################################
2> Install Multiple Instances Of Printer other than default instance

    # Open Config.xml file.
	
	# Change PrinterName, PortName, HardwareId, outputFilePath,outputFileName,paperSize,color parameters Values as needed.
	
		* Here PrinterName,PortName,HardwareId must be different for all instances of printer.
			i.e. For Second Instance Below Values are Okay.
		
	                    <Parameter Name="PrinterName" Value="PdfScribe2" />
            	            <Parameter Name="PortName" Value="PSCRIBE2" />
            	            <Parameter Name="HardwareId" Value="PDFScribe_Driver01012" />
		
			            <Parameter Name="outputFilePath" Value="D:\\PDF_SCRIBE\\Output" />			
			            <Parameter Name="outputFileName" Value="PdfDoc.pdf" />
			            <Parameter Name="paperSize" Value="A0" />
			            <Parameter Name="color" Value="Color" />

	# Save Config.xml File

	# Open Install_Instance.bat  file in any Text Editor(i.e. Notepad++)

	# Check TRANSFORMS=":instanceX" Parameter, And change X according to Instance number.(X >=2 And X<=10)

		e.x. --> Suppose You want to install Second Instance then change X to 2.(i.e. TRANSFORMS=":instance2")
	
	# check INSTALLFOLDER="" parameter, And Put installation path in string where you want to install, Which prefix installation path.
		
		e.x. --> Suppose You want to install second Instance on "C:\Program Files\PdfScribe2" path, then put it 
			like (INSTALLFOLDER="C:\Program Files\PdfScribe2")

	# Save the File And Run the Install_Instance.bat File.(See Important Notes Below)

		e.x. -->  Suppose You want to install Third Instance then change X to 3.(i.e. "TRANSFORMS=":instance3") and 
				change INSTALLFOLDER parameter to installation path.(i.e. INSTALLFOLDER="C:\Program Files\PdfScribe3")

	# Save the File And Run the Install_Instance.bat File.	


###################################################################################################################################

Important Notes:

# You must have to Install New Instance of printer to different Directory.

	# Suppose For Default Instance Installation Directory is "C:\ProgramFiles\PdfScribe\"	
	# For New Instance You have to set New Directory other than Default Instance Directory for Proper Working of Priner.
	# i.e You can set C:\ProgramFiles\PdfScribe\YourDesired_Name or C:\ProgramFiles\YourDesired_Name
	# But Not same Directory "C:\ProgramFiles\PdfScribe\".