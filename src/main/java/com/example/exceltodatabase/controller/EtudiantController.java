package com.example.exceltodatabase.controller;

import com.example.exceltodatabase.service.EtudiantService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class EtudiantController {
    @Autowired
    private EtudiantService etudiantService;

    @GetMapping()
    public ResponseEntity<String> handleFileUpload() {
        etudiantService.processExcelFile();
        return ResponseEntity.ok("File uploaded successfully");
    }
}
