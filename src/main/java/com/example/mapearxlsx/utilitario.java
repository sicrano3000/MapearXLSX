package com.example.mapearxlsx;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.Iterator;
import java.util.List;

class GerarCsvFromExcel {
    public static void main(String[] args) throws FileNotFoundException {
        InputStream in = new FileInputStream("/home/jhoncrespo/Downloads/mapeamentoFcu/FCU0054670_v6_CLN003TM_arquivo FLY.xlsx");
        createFile(in, 0);
    }

    public static void createFile(InputStream planilha, int sheet) {
        try {

            File file = new File(String.format("/home/jhoncrespo/Downloads/mapeamentoFcu/resultado-%d.csv", sheet));
            FileWriter fileWriter = new FileWriter(file, false);
            PrintWriter printWriter = new PrintWriter(fileWriter);


            XSSFWorkbook workbook = new XSSFWorkbook(planilha);
            Sheet sheetAt = workbook.getSheetAt(sheet);

            List<XSSFRow> lista = (List<XSSFRow>) getListFromIterator(sheetAt.iterator());
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter dataFormatter = new DataFormatter();
            dataFormatter.addFormat("m/d/yy", new java.text.SimpleDateFormat("yyyy-MM-dd"));

            for (XSSFRow linha : lista) {
                List<Cell> colunas = (List<Cell>) getListFromIterator(linha.cellIterator());
                colunas.forEach(coluna -> {
                            StringBuilder linhaArquivo = new StringBuilder();
                            switch (coluna.getCellType()) {
                                case STRING -> {
                                    linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(coluna.getStringCellValue());
                                }
                                case NUMERIC -> {
                                    if (DateUtil.isCellDateFormatted(coluna)) {
                                        String valorColuna = dataFormatter.formatCellValue(coluna, formulaEvaluator);
                                        linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(valorColuna);
                                    } else {
                                        linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(String.valueOf(BigDecimal.valueOf(coluna.getNumericCellValue())));
                                    }
                                }
                                case FORMULA -> {
                                    if (coluna.getCachedFormulaResultType() == CellType.NUMERIC) {
                                        linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(String.valueOf(BigDecimal.valueOf(coluna.getNumericCellValue())));
                                    } else if (coluna.getCachedFormulaResultType() == CellType.STRING) {
                                        linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(coluna.getStringCellValue());
                                    } else if (coluna.getCachedFormulaResultType() == CellType.BOOLEAN) {
                                        linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(String.valueOf(coluna.getBooleanCellValue()));
                                    } else {
                                        linhaArquivo.append(linha.getRowNum()).append(";").append(coluna.getColumnIndex()).append(";").append(coluna.getStringCellValue());
                                    }
                                }
                            }

                            if (StringUtils.isNotBlank(linhaArquivo)) {
                                printWriter.println(linhaArquivo);
                            }
                        }

                );
            }

            printWriter.close();

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

    public static List<?> getListFromIterator(Iterator<?> iterator) {
        return IteratorUtils.toList(iterator);
    }
}