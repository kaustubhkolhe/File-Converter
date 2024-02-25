package com.converter;

public class Main {

    public static void main(String[] args) {
        String userHome = System.getProperty("user.home");
        String basePath = userHome + "\\JIO Internship";

        Converter.xlsxToJson(basePath);
//        T2.jsonToCsv(basePath);
//        Converter.jsonToXlsx(basePath);
    }
}