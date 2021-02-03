package me.prj;

import org.apache.poi.hssf.usermodel.HSSFCell;

import javax.swing.*;
import javax.swing.border.Border;
import java.io.*;
import java.lang.reflect.Array;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

public class ExcelController {
    //    private String pathName;
//
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
//    public void mkExcel(){
//
//        workbook=new XSSFWorkbook();
//        sheet=workbook.createSheet();
//        //excel 헤더
//        row=sheet.createRow(rowIdx++);
//        cell=row.createCell(0);
//        cell.setCellValue("접수번호");
//        cell=row.createCell(1);
//        cell.setCellValue("신청인");
//        cell=row.createCell(2);
//        cell.setCellValue("수임담당");
//
//        //구분자는 원하는대로 바꿀수 있음
//        splitFileName(fileList);
//        mkFile();
//
//
//
//    }

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
    //    XSSFCellStyle styleCell(Enum<VerticalAlignment> vCenter,Enum<HorizontalAlignment> hCenter, Enum<BorderStyle> borderStyle){
//        XSSFCellStyle style;
//        style=workbook.createCellStyle();
//
//        style.setVerticalAlignment((VerticalAlignment) vCenter);
//        style.setAlignment((HorizontalAlignment) hCenter);
//        style.setBorderTop((BorderStyle) borderStyle);
//        style.setBorderBottom((BorderStyle) borderStyle);
//
//
//        return style;
//    }
    public void editFile() throws IOException {
        try{
            for(int i=0;i<fileList.size();i++){
                String caseNum=fileList.get(i).substring(0,9);
                String clientName=fileList.get(i).substring(9,12);

                String url=files[i].getAbsolutePath();
                fis=new FileInputStream(url);
                workbook=new XSSFWorkbook(fis);

//                XSSFFont fontBold=workbook.createFont();
//                fontBold.setFontHeightInPoints((short)12);
//                fontBold.setBold(true);
//                fontBold.setFontName("맑은고딕");
//
//                XSSFFont fontNormal=workbook.createFont();
//                fontNormal.setFontHeightInPoints((short)11);
//                fontNormal.setFontName("맑은고딕");
//
//                XSSFCellStyle styleCenter;
//                styleCenter=workbook.createCellStyle();
//                styleCenter.setAlignment(HorizontalAlignment.CENTER);
//                styleCenter.setVerticalAlignment(VerticalAlignment.CENTER);
//                styleCenter.setBorderTop(BorderStyle.MEDIUM);
//                styleCenter.setBorderBottom(BorderStyle.MEDIUM);


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
//                CellReference ref = new CellReference("I3");
//                XSSFRow row=sheet.getRow(ref.getRow());
//                cell=sheet.getRow(5).createCell(1);
//                cell.setCellFormula("i3");
//                style=styleCell(VerticalAlignment.CENTER,
//                        HorizontalAlignment.CENTER,
//                        BorderStyle.MEDIUM,
//                        BorderStyle.THIN,
//                        "맑은고딕",
//                        (short)11,
//                        false);
//                cell.setCellStyle(style);
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

//            for(String file_:fileList){
//                String caseNum=file_.substring(0,9);
//                String clientName=file_.substring(9,12);
//
//                String url=this.pathName+"\\"+file_;
//                fis=new FileInputStream(url);
//                workbook=new XSSFWorkbook(fis);
//
//                /*
//                * sheet #
//                * 1. 보고서 양식(재산) ... 접수번호, 신청인2
//                * 2. 보험가액 .. 접수번호(ref가 연락처로 되어 있긴 한데... 음..)
//                * 4. 총괄표 .. 피해항목번호가 접수번호인듯.
//                * */
//
//                //1. 보고서 양식(재산) ... 접수번호, 신청인(1개는 i3 참고한거라 설정 불필요)
//                sheet=workbook.getSheetAt(1);
//                sheet.getRow(2).createCell(1).setCellValue(caseNum);
//                sheet.getRow(2).createCell(8).setCellValue(clientName);
//
//                outStream=new FileOutputStream(url);
//                workbook.write(outStream);
//
//            }

        }catch(IOException ioe){
            ioe.printStackTrace();
        }finally {
            fis.close();
            outStream.close();
        }

    }
//    public void mkFile() {
//        String url=this.pathName+"\\"+this.fileName+".xlsx";
//        //윈도우에서는 역슬레시 써야함!!! 나중에 배포할때 변경
////        System.out.println(this.pathName+"/"+this.fileName+".xlsx");
//        File file=new File(url);
//        FileOutputStream fos=null;
//
//        try{
//            fos=new FileOutputStream(file);
//            workbook.write(fos);
//        }catch(FileNotFoundException fnf){
//            fnf.printStackTrace();
//            JOptionPane.showMessageDialog(null,"파일 생성에 실패했어요..."+url);
//        } catch (IOException e) {
//            e.printStackTrace();
//        } finally {
//            try {
//                if(workbook!=null) workbook.close();
//                if(fos!=null) fos.close();
//
//            } catch (IOException e) {
//                // TODO Auto-generated catch block
//                e.printStackTrace();
//            }
//        }
//    }

//    private void splitFileName(ArrayList<String> fileList) {
//        //ex) 001129_01김은옥(김찬우)-평가자료
//        for(String file:fileList){
//            String caseNum=file.substring(0,9);
//            String clientName=file.substring(9,12);
//            String managerName=file.substring(13,16);
//            //String []str=file.split("_");
//            row=sheet.createRow(rowIdx++);
//            cell=row.createCell(0);
//            cell.setCellValue(caseNum);
//            cell=row.createCell(1);
//            cell.setCellValue(clientName);
//            cell=row.createCell(2);
//            cell.setCellValue(managerName);
//
////            for(int i=0;i<str.length;i++) {
////                cell=row.createCell(i);
////                cell.setCellValue(str[i]);
////            }
//        }
//    }

//    cell=row.createCell(0);
}
