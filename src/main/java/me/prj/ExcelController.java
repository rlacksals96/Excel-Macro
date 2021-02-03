package me.prj;


import java.io.*;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;

public class ExcelController {

    private ArrayList<String> fileList;
    private int rowIdx=0;
    private File[] files;
    public ExcelController(File[] files,ArrayList<String> fileList){
        this.files=files;
        this.fileList=fileList;
    }
    XSSFWorkbook workbook;
    XSSFSheet sheet;

    FileOutputStream outStream=null;


    InputStream fis=null;


    XSSFCellStyle styleCell(Enum<VerticalAlignment> vCenter,Enum<HorizontalAlignment> hCenter, Enum<BorderStyle> borderTop,Enum<BorderStyle> borderBottom,String fontName,short fontSize,boolean isBold){
        XSSFCellStyle style;
        style=workbook.createCellStyle();

        style.setVerticalAlignment((VerticalAlignment) vCenter);
        style.setAlignment((HorizontalAlignment) hCenter);
        style.setBorderTop((BorderStyle) borderTop);
        style.setBorderBottom((BorderStyle) borderBottom);

        XSSFFont font= workbook.createFont();
        font.setFontHeightInPoints(fontSize);
        font.setFontName(fontName);
        font.setBold(isBold);

        style.setFont(font);

        return style;
    }

    public void editFile() throws IOException {
        try{
            for(int i=0;i<fileList.size();i++){
                String caseNum=fileList.get(i).substring(0,9);
                String clientName=fileList.get(i).substring(9,12);

                String url=files[i].getAbsolutePath();
                fis=new FileInputStream(url);
                workbook=new XSSFWorkbook(fis);


                /*
                 * sheet #
                 * 1. 보고서 양식(재산) ... 접수번호, 신청인2
                 * 2. 보험가액 .. 접수번호(ref가 연락처로 되어 있긴 한데... 음..)
                 * 4. 총괄표 .. 피해항목번호가 접수번호인듯.
                 * 5. 내역서 .. "공사명: " + 이름 + " 보수공사"
                 * */

                //1. 보고서 양식(재산) ... 접수번호, 신청인(1개는 i3 참고한거라 설정 불필요)

                //접수번호
                sheet=workbook.getSheetAt(1);


                XSSFCell cell=sheet.getRow(2).createCell(1);
                cell.setCellValue(caseNum);
                XSSFCellStyle style=styleCell(VerticalAlignment.CENTER,
                        HorizontalAlignment.CENTER,
                        BorderStyle.MEDIUM,
                        BorderStyle.MEDIUM,
                        "맑은고딕",
                        (short)11,
                        false);
                cell.setCellStyle(style);


                //신청인
                cell=sheet.getRow(2).createCell(8);
                cell.setCellValue(clientName);
                style=styleCell(VerticalAlignment.CENTER,
                        HorizontalAlignment.CENTER,
                        BorderStyle.MEDIUM,
                        BorderStyle.MEDIUM,
                        "맑은고딕",
                        (short)12,
                        true);
                cell.setCellStyle(style);

                //일반사항.신청인
                //별도 처리 없이 계산만 다시 해주면 됨.
                workbook.setForceFormulaRecalculation(true);

//          4. 총괄표 .. 피해항목번호가 접수번호인듯.
                sheet=workbook.getSheetAt(4);
                cell=sheet.getRow(1).createCell(5);
                cell.setCellValue(caseNum);
                style=styleCell(VerticalAlignment.CENTER,
                        HorizontalAlignment.CENTER,
                        BorderStyle.NONE,
                        BorderStyle.NONE,
                        "맑은고딕",
                        (short)11,
                        true);
                cell.setCellStyle(style);


//          5. 내역서 .. "공사명: " + 이름 + " 보수공사"
                sheet=workbook.getSheetAt(5);
                cell=sheet.getRow(1).createCell(0);
                cell.setCellValue("공사명: "+clientName+" 보수공사");
                style=styleCell(VerticalAlignment.CENTER,HorizontalAlignment.LEFT,BorderStyle.NONE,BorderStyle.NONE,"맑은고딕",(short)10,false);
                cell.setCellStyle(style);


                outStream=new FileOutputStream(url);
                workbook.write(outStream);
            }



        }catch(IOException ioe){
            ioe.printStackTrace();
        }finally {
            fis.close();
            outStream.close();
        }

    }

}
