package com.reporte.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class ReporteService {

    public List<Object> listObjExcel = new ArrayList<>();

    public boolean generateExcel(String fileName){
        this.fillDataExcel();
        try {

            SXSSFWorkbook wb = new SXSSFWorkbook(1);
            SXSSFSheet sheet = wb.createSheet();
            Row nRow = null;
            Cell nCell = null;

            // Generar la cabecera
            Object[] objHead = {"NOMBRES","APELLIDOS","DIRECCION"};
            nRow = sheet.createRow(0);
            for (int i = 0; i < objHead.length; i++) {
                nCell = nRow.createCell(i);
                nCell.setCellValue(objHead[i].toString());
            }

            //Generado el cuerpo del excel
            Iterator<Object> it = listObjExcel.iterator();
            int pageRowNo = 1;
            while (it.hasNext()){
                Object[] objExcelBody = (Object[]) it.next();
                nRow = sheet.createRow(pageRowNo++);
                for (int i = 0; i < objExcelBody.length; i++) {
                    nCell = nRow.createCell(i);
                    nCell.setCellValue(objExcelBody[i].toString());
                }
                it.remove();
            }

            FileOutputStream fileOutputStream = new FileOutputStream(fileName);
            wb.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();
            wb.dispose();
            return true;
        }catch (Exception e){
            return false;
        }
    }


    private void fillDataExcel(){
        for (int i = 0; i <= 500000; i++) {
            Object[] objBody = {"Nombre "+i,"Apellido "+i,"Calle "+i};
            listObjExcel.add(objBody);
        }
    }

}
