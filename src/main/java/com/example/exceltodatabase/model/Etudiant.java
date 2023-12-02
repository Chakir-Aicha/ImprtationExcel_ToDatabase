package com.example.exceltodatabase.model;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import lombok.*;

import java.util.Date;

@Entity
@Data
@NoArgsConstructor
@AllArgsConstructor
@Setter
@Getter
@ToString
public class Etudiant {
    @Id
    private String CNE;
    private String nom;
    private String prenom;
    private Date date_naissance;
    private float note;
    private String mention;
}
