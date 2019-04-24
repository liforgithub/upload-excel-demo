package com.lxy.controller;

import com.lxy.service.ResolveExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/resolve")
public class ExcelController {

    @Autowired
    private ResolveExcelService resolveExcelService;

    /**
     * 文件上传
     */
    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    @ResponseBody
    public Object uploadExcel(@RequestParam("file") MultipartFile excelFile) throws Exception {
        resolveExcelService.resolveExcel(excelFile);
        return "success";
    }
}
