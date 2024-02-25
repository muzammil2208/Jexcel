package excellib;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excellib {
    String path;
    public  HashMap<String,String> result;
    public  XSSFWorkbook workbook;
    public excellib()
    {
        this.path=System.getProperty("user.dir") + "/TestData/TestData.xlsx";
    }
    public excellib(String path)
    {

        this.path=path;
    }

    
    public  String getTestData(String key)
    {
        return result.get(key);
    }
    public ArrayList<String> getSheetNames()
    {
        ArrayList<String> names=new ArrayList<>();
        try
        {
            InputStream file=new FileInputStream(path);
            workbook=new XSSFWorkbook(file);
            int count=workbook.getNumberOfSheets();
            if(count<=0)
                throw new Exception("no sheets present in the workbook");
            for(int i=0;i<count;i++)
            {
                names.add(workbook.getSheetName(i));
            }
            file.close();
        }
        catch(Exception e)
        {
            System.out.println("Error while handling excel sheet:");
            System.out.println(e.getMessage());
        }
        return names;


    }
    public int getColumnsCount(String sheetName)
    {
        int result=0;
        try
        {
            FileInputStream file=new FileInputStream(path);
            workbook=new XSSFWorkbook(file);

            Sheet sheet=workbook.getSheet(sheetName);
            if(sheet==null) {
                throw new Exception("no such sheet with name:" + sheetName + " exist");
            }
            Row row=sheet.getRow(0);
            for(Cell cell:row)
                result++;
        }
        catch(Exception e)
        {
            System.out.println("Error while getting column counts:");
            System.out.println(e.getMessage());
        }
        return result;
    }
    public int getRowsCount(String sheetName)
    {
        int result=0;
        try
        {
            FileInputStream file=new FileInputStream(path);
            workbook=new XSSFWorkbook(file);

            Sheet sheet=workbook.getSheet(sheetName);
            if(sheet==null) {
                throw new Exception("no such sheet with name:" + sheetName + " exist");
            }
            for(Row row:sheet)
            {
                result++;
            }
        }

        catch(Exception e)
        {
            System.out.println("Error while getting row counts:");
            System.out.println(e.getMessage());
        }
        return result;
    }
    public int getSheetsCount()
    {
        int result=0;
        try
        {
            FileInputStream file=new FileInputStream(path);
            workbook=new XSSFWorkbook(file);
            result=workbook.getNumberOfSheets();
            file.close();
        }
        catch(Exception e)
        {
            System.out.println("Error while retrieving data from test data file:");
            System.out.println(e.getMessage());
        }
        return result;
    }
    public  String getCellValue(Cell cell)
    {
        DataFormatter dataFormatter = new DataFormatter();
        return dataFormatter.formatCellValue(cell);
    }
    public void setTestData(String sheetName,String testCaseName)
    {
        result=new HashMap<>();
        ArrayList<String> keys=new ArrayList<>();
        ArrayList<String> values=new ArrayList<>();
        try {


            InputStream file = new FileInputStream(path);
            workbook = new XSSFWorkbook(file);
            Sheet sheet=workbook.getSheet(sheetName);

            Row firstRow=sheet.getRow(0);
            //creating keys of hashmaps
            for(Cell cell:firstRow)
            {
                keys.add(getCellValue(cell));
            }



            //intializing data for selected testcase
            for(Row row:sheet)
            {
                Cell firstCell=row.getCell(0);
                String tcname=getCellValue(firstCell);

                if(tcname.equalsIgnoreCase(testCaseName))
                {
                    for(Cell cell:row)
                    {

                        values.add(getCellValue(cell));
                    }

                }

            }

            //joining both keys and values into hashmap
            for(int i=0;i<keys.size();i++)
            {
                result.put(keys.get(i),values.get(i));
            }
            workbook.close();
            file.close();
        }
        catch(Exception e)
        {
            System.out.println("error while retrieving data from excel/n more detail"+e.getMessage());
        }


    }

    public String getTestDataByCoordinates(String sheetName,int rowCount,int columnCount)
    {
        String result="";
        try
        {
            FileInputStream file=new FileInputStream(path);
            workbook=new XSSFWorkbook(file);
            Sheet sheet=workbook.getSheet(sheetName);
            Row row=sheet.getRow(rowCount-1);
            result=getCellValue(row.getCell(columnCount-1));
            file.close();
        }
        catch(Exception e)
        {
            System.out.println("Error while getting data by coordinates:");
            System.out.println(e.getMessage());
        }
        return result;
    }
}


