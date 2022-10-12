package com.reporte.Controller;

import com.reporte.service.ReporteService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api")
@CrossOrigin
public class ApiController {

    @Autowired
    ReporteService reporteService;

    @GetMapping("/reporte")
    public String getReporte(){
        if(reporteService.generateExcel("C:\\www\\reporteMasivo.xlsx"))
            return  "Reporte generado correctamente";
        else
            return "Ha ocurrido un error en la generacion del excel";
    }

}
