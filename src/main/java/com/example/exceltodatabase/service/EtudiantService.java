package com.example.exceltodatabase.service;

import com.example.exceltodatabase.model.Error;
import com.example.exceltodatabase.model.Etudiant;
import com.example.exceltodatabase.repository.EtudiantRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.exception.ConstraintViolationException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

@Service
public class EtudiantService {
    @Autowired
    private EtudiantRepository etudiantRepository;

    @Autowired
    private ResourceLoader resourceLoader;

    public void processExcelFile() {
        try {
            // Charger le fichier Excel depuis les ressources
            Resource resource = new ClassPathResource("test.xlsx");
            InputStream inputStream = resource.getInputStream();

            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            //list errors
            ArrayList<Error> errors = new ArrayList<>();

            for (Row row : sheet) {
                // Ignorer la ligne d'en-tête contient les attribus
                if (row.getRowNum() == 0) {
                    continue;
                }

                Etudiant etudiant = new Etudiant();

                // Gérer les erreurs de colonne vide
                if (row.getCell(0) == null || row.getCell(1) == null || row.getCell(2) == null) {
                    errors.add(new Error(row.getRowNum() ,"Une ou plusieurs colonnes sont vides."));
                    continue;
                }
                etudiant.setCNE(row.getCell(0).getStringCellValue());
                etudiant.setNom(row.getCell(1).getStringCellValue());
                etudiant.setPrenom(row.getCell(2).getStringCellValue());
                etudiant.setDate_naissance(row.getCell(3).getDateCellValue());
                etudiant.setNote((float) row.getCell(4).getNumericCellValue());
                etudiant.setMention(row.getCell(5).getStringCellValue());
                //gérer l'erreur de duplication de la clé primaire
                Etudiant existingEtudiant = etudiantRepository.findByCNE(etudiant.getCNE());
                if (existingEtudiant != null) {
                    errors.add(new Error(row.getRowNum() , "Duplication de clé primaire - CNE " + etudiant.getCNE() + " déjà existant."));
                    continue;
                }
                try {
                    etudiantRepository.save(etudiant);
                } catch (DataIntegrityViolationException e) {
                    if (e.getCause() instanceof ConstraintViolationException) {
                        ConstraintViolationException cause = (ConstraintViolationException) e.getCause();
                        if (cause.getConstraintName().contains("foreign key constraint name")) {
                            // Gérer l'erreur de clé étrangère
                            errors.add(new Error(row.getRowNum() + 1, " Erreur de clé étrangère - " + e.getMessage()));
                        } else if(cause.getConstraintName().contains("primary key constraint name")){
                            // Autre erreur de contrainte d'intégrité
                            errors.add(new Error(row.getRowNum(), " Erreur de duplication du clé étrangère - " + e.getMessage()));
                        }
                    } else {
                        // Autre type d'erreur lors de l'enregistrement
                        errors.add(new Error(row.getRowNum() ," Erreur lors de l'enregistrement en base de données - " + e.getMessage()));
                    }
                }
            }

            workbook.close();

            if (!errors.isEmpty()) {
                generateErrorExcelFile(errors,sheet);
            }
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Error processing Excel file", e);
        }
    }
    // methode pour générer le fichier Excel avec les erreurs
    private void generateErrorExcelFile(ArrayList<Error> errors, Sheet originalSheet) {
        // La classe XSSFWorkbook est spécifiquement conçue pour traiter le format XLSX.
        // Crée un nouveau classeur Excel au format XLSX
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Erreurs");
        // Créer une ligne d'en-tête
        Row headerRow = sheet.createRow(0);
    // Ajouter les en-têtes des colonnes
        headerRow.createCell(0).setCellValue("CNE");
        headerRow.createCell(1).setCellValue("Nom");
        headerRow.createCell(2).setCellValue("Prenom");
        headerRow.createCell(3).setCellValue("Date de Naissance");
        headerRow.createCell(4).setCellValue("Note");
        headerRow.createCell(5).setCellValue("Mention");
        headerRow.createCell(6).setCellValue("Erreur");
    // Ajouter les erreurs à la feuille
        for (int i = 0; i < errors.size(); i++) {
            Row row = sheet.createRow(i + 1);
            // Si vous avez la référence à la ligne d'origine où l'erreur s'est produite
            // Copier les données de la ligne originale dans la feuille d'erreurs
            Row originalRow = originalSheet.getRow(errors.get(i).getRowNum());
            if (originalRow != null) {
                for (int j = 0; j <= 5; j++) {
                    Cell originalCell = originalRow.getCell(j);
                    if (originalCell != null) {
                        row.createCell(j).setCellValue(getCellValueAsString(originalCell));
                    }
                }
            }
            // Ajouter le message d'erreur à la dernière colonne
            row.createCell(6).setCellValue(errors.get(i).getMessage());
        }
        // Enregistrer le classeur Excel dans un fichier
        try (FileOutputStream fileOut = new FileOutputStream("errors.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("Error generating error Excel file", e);
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private String getCellValueAsString(Cell cell) {
        // Méthode utilitaire pour obtenir la valeur de la cellule sous forme de chaîne
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                // Formater la date si la cellule contient une date
                return cell.getDateCellValue().toString();
            } else {
                return Double.toString(cell.getNumericCellValue());
            }
        }
        if (cell.getCellType() == CellType.BOOLEAN) {
            return Boolean.toString(cell.getBooleanCellValue());
        }
        return null;
    }
}
