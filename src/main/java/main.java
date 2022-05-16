import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class main {
    public static void main(String[] args) throws IOException {
        System.out.println("excel to json");
        //read excel file
        String filepath="/home/gp/Documents/file.xlsx";
        FileInputStream fis = new FileInputStream(filepath);
        XSSFWorkbook book= new XSSFWorkbook(fis);
        List<Sheet> sheetList=new ArrayList<>();
        for (int i=0;i<book.getNumberOfSheets();i++){
            sheetList.add(book.getSheetAt(i));
        }
        int i=1;
//        for (Sheet sheet :
//                sheetList) {
//            System.out.println(""+i+" "+sheet.getSheetName());
//            i++;
//        }
        String sheetName=book.getSheetName(68);
        System.out.println("Sheet Name : "+sheetName);
      Sheet sheet=  book.getSheet(sheetName);
      Iterator<Row> iterator=sheet.iterator();
//      while (iterator.hasNext()){
//          Row row=iterator.next();
//          Iterator<Cell> cellIterator=row.iterator();
//          while (cellIterator.hasNext()){
//              Cell cell=cellIterator.next();
//              System.out.println(cell.getStringCellValue());
//          }
//      }
        List<String> columns=new ArrayList<>();
      Row row=  sheet.getRow(1);
        System.out.println(sheet.getLastRowNum());
      Iterator<Cell> iterator1=row.iterator();
        System.out.println("Coloumn names are below...\n");
      while (iterator1.hasNext()){
          Cell c= iterator1.next();
          columns.add(c.getStringCellValue());
      }
      List<CellRangeAddress> l=sheet.getMergedRegions();
     //tabs
        for(int j=1;j<sheet.getLastRowNum()+1;j++){
            Row row3=sheet.getRow(j);
            if(row3!=null) {
            Cell c=row3.getCell(2);
String itemname="";
    CellType ct = c.getCellType();
    switch (ct) {
        case BLANK:
            System.out.println(c.getRowIndex() + " " + ct);
            break;
        case STRING:
            String value=c.getStringCellValue();
            //System.out.println(c.getRowIndex() + " " + c.getStringCellValue());
            if(value.equals("TAB")){

            }else{
itemname=value;
                int index=c.getRowIndex()+1;
                int height=0;
                c=sheet.getRow(index).getCell(2);
                if(c!=null) {
                    while (c.getCellType() == CellType.BLANK) {
                        if (sheet.getRow(index) != null) {
                            c = sheet.getRow(index).getCell(2);
                            index++;
                            height++;
                        } else {
                            index++;
                        }
                    }
                }
                //j=index;
                System.out.println(index+" "+itemname+" "+height);

            }
            break;
    }
}
        }
    }
}
