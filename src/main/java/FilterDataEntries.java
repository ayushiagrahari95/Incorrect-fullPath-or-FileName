import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import org.apache.commons.io.FileUtils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class FilterDataEntries {
	File rootDirectory;
	int counter=0;
	File existingFullPath;
	
	public static void main(String args[])throws IOException{
		
		FilterDataEntries data=new FilterDataEntries();
		/*data.rootDirectory=new File(args[0]);*/
		File rootDir=new File(args[0]);
		data.rootDirectory=new File(rootDir.getAbsolutePath());/*storing the absolute path of root directory*/
		FileUtils.deleteDirectory( new File(data.rootDirectory,"Incorrect Entries"));/*delete the folder for running the jar next tym*/
		data.displayDirectoryContents(data.rootDirectory);	
	}
	
	/*recursive method to check if the contents of a directory is a file or folder and performing the required action*/
	void displayDirectoryContents(File dir) throws IOException{
		File[] files = dir.listFiles();/*listing of files and directories inside a diectory	*/
		/*iterating over each file and directory inside a directory*/
		for (File file : files)
		{/*checking for the existence of a directory inside a directory*/
		if	(file.isDirectory()){
			System.out.println("Directory found:"+file.getName());/*if directory is found inside a directory,recursively check the contents of the directory*/
			displayDirectoryContents(file);
		}
		
		/*checking for the existence of excel files in a directory*/
		else if(file.isFile() && (file.getName().toLowerCase().endsWith(".xlsx") || file.getName().toLowerCase().endsWith(".xls")))
		{
		counter++;
		String fileNameWithExtension=file.getName();
		String fileNameWithoutExtension=fileNameWithExtension.substring(0,fileNameWithExtension.lastIndexOf("."));/*trimming the extension of excel file names*/
		String textFileNameWithExtension=fileNameWithoutExtension+".txt";
		File folder=new File(rootDirectory,"Incorrect Entries");
		File textFiles=new File(folder,textFileNameWithExtension);
		String excelPath=file.getParent();
		System.out.println("Excel Path"+excelPath);
		System.out.println("Exccel file:"+fileNameWithExtension);
		System.out.println("Exccel file name without extension:"+fileNameWithoutExtension);
		try{
			/*reading excel files*/
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook book = new XSSFWorkbook(fis);
			XSSFSheet sheet = book.getSheetAt(0);
			System.out.println("No. of Sheets"+book.getNumberOfSheets()) ;
			/*iterating the rows of the excel file*/
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
				XSSFRow row = sheet.getRow(rowIndex);
				/*for each row,iterate the columns/ */
				for (int colIndex = 0; colIndex <= 1; colIndex++){
					XSSFCell cell=row.getCell(colIndex);
					if (colIndex==0){/*get the first cell of an excel file*/
						System.out.println("Column:"+colIndex);
						existingFullPath = new File(excelPath,cell.getStringCellValue());
						if (existingFullPath.exists())/*if full path exists check for file names*/
						{
						System.out.println("Fullpath"+existingFullPath+"exists,checking for file names");
						checkFileNames(existingFullPath,row,textFiles,folder);	
						}
						else
						{	/*if fullpath does not exists*/
							System.out.println("Non Existing Full Path"+existingFullPath);
							/*if folder doesnot exists,create folder,create text file if not already present and append incorrect entries to text files*/
							if(!folder.exists())
							{
								createFolder(folder);
							}
							/*if folder exists already,no need to create new folder.Check if the text files are created or not*/
							
							/*if text file does not exists,create text file and append data to the text file*/
							if(!textFiles.exists())
							{
								createTextFiles(textFiles);
								appendFullPath(row,cell,textFiles);
							}
							
							/*if text file already exists,just append data to the text file*/
							else{
								System.out.println("Text file "+textFileNameWithExtension+"exists,writing data to text files");
								appendFullPath(row,cell,textFiles);
							}
						}
					}
				}
			}
			book.close();
		}
		catch(FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("Number of excel files"+" "+counter);
		}
	}
		}
	
	/*recursive method that checks for the existence of the file name and then taking the appropriate action*/
	void checkFileNames(File ExistingFullPath,XSSFRow Row,File TextFiles,File Folder) throws IOException{
		int IncreementedColumnIndex=1;
		XSSFCell Cell=Row.getCell(IncreementedColumnIndex);
		File existingFiles=new File(ExistingFullPath,Cell.getStringCellValue());
		
		if(existingFiles.exists())/*if filename and full path exists do nothing*/
		{
			System.out.println("Existing FullPath:"+ExistingFullPath+'\t'+"Existing Files:"+existingFiles);
		}
		
		else{/*if filename does not exists*/
			System.out.println("Existing FullPath:"+ExistingFullPath+'\t'+"Non Existing Files:"+existingFiles);
			/*if folder doesnot exists,create folder,create text file if not already present and append incorrect entries to text files*/
			if(!Folder.exists())
			{
				createFolder(Folder);
			}
			/*if folder exists already,no need to create new folder.Check if the text files are created or not*/
			
			/*if text file does not exists,create text file and append data to the text file*/
			if(!TextFiles.exists())
			{
				createTextFiles(TextFiles);
				appendFileName(Row,TextFiles,Cell);
			}
			
			/*if text file already exists,just append data to the text file*/
			else{
				System.out.println("Text file "+TextFiles.getName()+"exists,writing data to text files");
				appendFileName(Row,TextFiles,Cell);
			}
		}
	}

	/*method to create folder if not exists*/
	void createFolder(File createFolder){
		createFolder.mkdir();
		System.out.println("Folder created");
	}
	
	void createTextFiles(File TextFiles) throws IOException{
		TextFiles.createNewFile();
		System.out.println("Text file Created:"+TextFiles.getName());
		FileWriter writtenToTextFiles = new FileWriter(TextFiles, true);
		BufferedWriter bw = new BufferedWriter(writtenToTextFiles);
		bw.write("Fullpath"+'\t'+'\t'+'\t'+'\t'+"Filename"+'\n');
		bw.close();
	}
	
	/*method to write the incorrect fullpath and filename in case of incorrect filename*/
	void appendFileName(XSSFRow rows,File TextFiless,XSSFCell cells) throws IOException{
		/*in case of incorrect file name,decreement column index for fullpath and write full path and file name to the text filename*/
		int decreementedColumnIndex=0;
		XSSFCell decreementedColIndex = rows.getCell(decreementedColumnIndex);
		FileWriter writtenToTextFiles = new FileWriter(TextFiless, true);
		BufferedWriter bw = new BufferedWriter(writtenToTextFiles);
		bw.write(decreementedColIndex.getStringCellValue() + '\t' + cells.getStringCellValue() + '\n');
		bw.close();
		System.out.println("incorrect fullpath alongwith file name written to the text file");

	}

	/*method to write the incorrect fullpath and filename in case of incorrect full path*/
	void appendFullPath(XSSFRow rows,XSSFCell cells,File TextFiles) throws IOException{
		/*in case of incorrect file name,increement column index for filename and write full path and file name to the text filename*/
		int increementedColumnIndex=1;
		XSSFCell increementedColIndex = rows.getCell(increementedColumnIndex);
		System.out.println("Non Existing Full path:"+existingFullPath);
		FileWriter writtenToTextFiles = new FileWriter(TextFiles, true);
		BufferedWriter bw = new BufferedWriter(writtenToTextFiles);
		bw.write(cells.getStringCellValue()+"\t"+increementedColIndex.getStringCellValue() + '\n');
		bw.close();
		System.out.println("incorrect fullpath alongwith file name written to the text file");
	}
}
