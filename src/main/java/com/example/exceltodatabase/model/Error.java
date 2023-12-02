package com.example.exceltodatabase.model;

public class Error {
    private int rowNum;
    private String message;

    public Error(int rowNum, String message) {
        this.rowNum = rowNum;
        this.message = message;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    @Override
    public String toString() {
        return "Error{" +
                "rowNum=" + rowNum +
                ", message='" + message + '\'' +
                '}';
    }
}
