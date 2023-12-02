package com.example.exceltodatabase.repository;

import com.example.exceltodatabase.model.Etudiant;
import org.springframework.data.jpa.repository.JpaRepository;

public interface EtudiantRepository extends JpaRepository<Etudiant,String> {
    Etudiant findByCNE(String cne);
}
