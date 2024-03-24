package com.autoworld.excelreader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

public class ReadExcel {
    private static ThreadLocal<InputStream> fis=new ThreadLocal<>();
    private static ThreadLocal<XSSFWorkbook> xssfWorkbook=new ThreadLocal<>();
    private static ThreadLocal<XSSFSheet> xssfSheet=new ThreadLocal<>();
    private static DataFormatter dataFormatter=new DataFormatter();
    private static FormulaEvaluator formulaEvaluator;
    private static Logger logger=Logger.getLogger(ReadExcel.class.getName());

    private ReadExcel(){}

    private static void setup(String excelName,String sheetName) throws IOException {
        InputStream inputStreamXls=ReadExcel.class.getResourceAsStream("/data/"+excelName+".xls");
        InputStream inputStreamXlsx=ReadExcel.class.getResourceAsStream("/data/"+excelName+".xlsx");
        if(inputStreamXls!=null){
            fis.set(inputStreamXls);
        }else{
            if(inputStreamXlsx==null){
                logger.warning("File not found");
            }
            fis.set(inputStreamXlsx);
        }
        xssfWorkbook.set(new XSSFWorkbook(fis.get()));
        xssfSheet.set(xssfWorkbook.get().getSheet(sheetName));
        xssfWorkbook.get().close();

    }

    public static ArrayList<Map<String,String>> readData(String fileName, String sheetName){
        System.out.println("Before read data: "+Thread.currentThread().getId());
        ArrayList<Map<String,String>> excelData=new ArrayList<>();

        try{
            setup(fileName,sheetName);
            int rowNums=xssfSheet.get().getLastRowNum();
            for(int i=1;i<=rowNums;i++){
                Map<String,String> mapData=getMapDataFromRow(i);
                excelData.add(mapData);
            }

        }catch (Exception e){
            logger.warning(e.getMessage());
        }finally {
            IOUtils.closeQuietly((Closeable) fis.get());
        }
        return excelData;
    }

    private static Map<String, String> getMapDataFromRow(int ronNum) {
        String[] headerData=getDataFromRow(0);
        String[] rowData=getDataFromRow(ronNum);

        Map<String,String> result=new HashMap<>();
        for(int i=0;i< headerData.length;i++){
            if(i>=rowData.length){
                result.put(headerData[i],"");
            }else{
                result.put(headerData[i],rowData[i]);
            }
        }
        return result;
    }

    private static String[] getDataFromRow(int rowNum) {
        int cellNum=xssfSheet.get().getRow(rowNum).getLastCellNum();
        String[] rowData=new String[cellNum];

        for(int i=0;i<cellNum;i++){
            rowData[i]=getStringValue(xssfSheet.get().getRow(rowNum).getCell(i));
        }
        return rowData;
    }

    private static String getStringValue(Cell cell) {
        formulaEvaluator=xssfSheet.get().getWorkbook().getCreationHelper().createFormulaEvaluator();
        if(cell!=null){
        if(cell.getCellType().equals(CellType.BOOLEAN)){
            return String.valueOf(cell.getBooleanCellValue());
        }else if(cell.getCellType().equals(CellType.NUMERIC)){
            return dataFormatter.formatCellValue(cell);
        }else if(cell.getCellType().equals(CellType.STRING)){
            return cell.getRichStringCellValue().getString();
        }else if(cell.getCellType().equals(CellType.FORMULA)){
            return formulaEvaluator.evaluate(cell).formatAsString();
        }
        }
        return "";
    }
}
