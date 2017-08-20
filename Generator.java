/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.seta.java.javafx.util.tables.generation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author moslim
 */
public class Generator
{

    public class GenerationStrategy
    {

        private boolean fxml;
        private String modelName;
        private String viewName;

        public GenerationStrategy(boolean fxml, String modelName, String viewName)
        {
            this.fxml = fxml;
            this.modelName = modelName;
            this.viewName = viewName;
        }

    }

    private static List<String> getHeaders(File excelFile) throws IOException
    {
        FileInputStream inputStream = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet spreadsheet = workbook.getSheetAt(0);
        int numberOfColumns = spreadsheet.getRow(0).getPhysicalNumberOfCells();
        XSSFRow row = spreadsheet.getRow(0);
        List<String> headers = new LinkedList<>();
        for (int i = 0; i < numberOfColumns; i++)
        {
            headers.add(getCellValue(row.getCell(i)));
        }
        return headers;
    }

    private static String generateClassDeclaration(String klass)
    {
        return "public class " + klass + "\n{";
    }

    private static String getLoaderMethodHeader(String modelKlassName)
    {
        return String.format("    public static List<%s> load(File file)\n"
                + "    {\n"
                + "        List<%s> models = new LinkedList<>();\n"
                + "\n"
                + "        try\n"
                + "        {\n"
                + "            FileInputStream inputStream = new FileInputStream(file);\n"
                + "            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);\n"
                + "            XSSFSheet spreadsheet = workbook.getSheetAt(0);\n"
                + "\n"
                + "            Iterator<Row> rowIterator = spreadsheet.iterator();\n"
                + "            XSSFRow row = (XSSFRow) rowIterator.next();\n"
                + "            %s item;\n"
                + "            XSSFCell cell;\n"
                + "            while (rowIterator.hasNext())\n"
                + "            {\n"
                + "                row = (XSSFRow) rowIterator.next();\n"
                + "                Iterator<Cell> celIterator = row.iterator();\n"
                + "                item = new %s();\n"
                + "                cell = (XSSFCell) celIterator.next();\n", modelKlassName, modelKlassName, modelKlassName, modelKlassName);
    }

    private static String getLoaderMethodTail()
    {
        return "            }\n"
                + "        } catch (Exception e)\n"
                + "        {\n"
                + "            e.printStackTrace();\n"
                + "        }\n"
                + "        return models;\n"
                + "    }";
    }

    public static class Generation
    {

        public String view;
        public String model;

        public Generation(String view, String model)
        {
            this.view = view;
            this.model = model;
        }

    }

    /**
     *
     * @param pakage package name of the generated classes
     * @param modelName name of the class that represents table model
     * @param viewName name of the class that represents table view
     * @param srcFile the Excel file
     * @param anootated if you want to use JavaFX @FXML annotations or not
     * @return
     */
    public static Generation generate(String pakage, String modelName, String viewName, File srcFile, boolean anootated)
    {

        StringBuilder model = new StringBuilder();
        StringBuilder view = new StringBuilder();
        try
        {
            List<String> headers = getHeaders(srcFile);
            model.append("package ").append(pakage).append(";").append(System.lineSeparator());
            model.append("import java.io.File;");
            model.append("import java.io.FileInputStream;\n"
                    + "import java.util.Iterator;\n"
                    + "import java.util.LinkedList;\n"
                    + "import java.util.List;\n"
                    + "import org.apache.poi.ss.usermodel.Cell;\n"
                    + "import org.apache.poi.ss.usermodel.Row;\n"
                    + "import org.apache.poi.xssf.usermodel.XSSFCell;\n"
                    + "import org.apache.poi.xssf.usermodel.XSSFRow;\n"
                    + "import org.apache.poi.xssf.usermodel.XSSFSheet;\n"
                    + "import org.apache.poi.xssf.usermodel.XSSFWorkbook;\n");

            model.append(generateClassDeclaration(modelName)).append("\n");
            model.append("private static String getCellValue(XSSFCell cell)\n"
                    + "    {\n"
                    + "        switch (cell.getCellTypeEnum())\n"
                    + "        {\n"
                    + "            case STRING:\n"
                    + "                return cell.getStringCellValue().trim();\n"
                    + "            case BOOLEAN:\n"
                    + "                return cell.getBooleanCellValue() + \"\";\n"
                    + "            case ERROR:\n"
                    + "                return cell.getErrorCellString() + \"\";\n"
                    + "            case NUMERIC:\n"
                    + "                return cell.getNumericCellValue() + \"\";\n"
                    + "            case BLANK:\n"
                    + "                return \"\";\n"
                    + "            case FORMULA:\n"
                    + "                return cell.getCellFormula();\n"
                    + "            default:\n"
                    + "                return \"\";\n"
                    + "        }\n"
                    + "    }");
            view.append("package ").append(pakage).append(";").append(System.lineSeparator());
            view.append("import javafx.scene.control.TableColumn;\n"
                    + "import javafx.scene.control.TableView;\n"
                    + "import javafx.fxml.FXML;\n"
                    + "import javafx.scene.control.cell.PropertyValueFactory;\n");
            view.append(generateClassDeclaration(viewName)).append("\n");
            if (anootated)
            {
                view.append("    @FXML\n");
            }
            view.append(String.format("    private TableView<%s> table;", modelName)).append("\n");
            StringBuilder loaderMethod = new StringBuilder();
            StringBuilder viewConstructor = new StringBuilder();
            viewConstructor.append(String.format("    public %s()\n    {\n", viewName));
            viewConstructor.append(String.format("        this.table = new TableView<>();\n"));
            loaderMethod.append(getLoaderMethodHeader(modelName));
            headers.forEach(name ->
            {
                String columnName = name + "Column";
                model.append(declareField(name, false));
                view.append(declareParamerized(columnName, modelName, anootated));

                loaderMethod.append(String.format("                item.set%s(getCellValue(cell));\n", generateName(name)));

                viewConstructor.append(String.format("        %s = new TableColumn<>();\n", columnName));
                viewConstructor.append(String.format("        %s.setCellValueFactory(new PropertyValueFactory<>(\"%s\"));\n", columnName, name));
                viewConstructor.append(String.format("        table.getColumns().add(%s);\n", columnName));

            });
            loaderMethod.append("                models.add(item);\n");
            loaderMethod.append(getLoaderMethodTail());
            headers.forEach(name ->
            {
                model.append(declareAccessor(name));
            });
            model.append(loaderMethod.toString()).append(System.lineSeparator()).append("}");
            viewConstructor.append("    }");
            view.append(viewConstructor.toString());
            view.append(String.format(" public TableView<Task> getTable()\n"
                    + "    {\n"
                    + "        return table;\n"
                    + "    }\n", modelName));
            view.append("}");
        } catch (Exception e)
        {
            e.printStackTrace();
        }
        return new Generation(view.toString(), model.toString());
    }

    private static String declareField(String name, boolean annotated)
    {
        if (annotated)
        {
            return String.format("    @FXML\n    private String %s;\n", name);
        }
        return String.format("    private String %s;\n", name);
    }

    private static String declareParamerized(String name, String parameter, boolean annotated)
    {
        if (annotated)
        {
            return String.format("    @FXML\n    private TableColumn<%s,String> %s;\n", parameter, name);
        }
        return String.format("    private TableColumn<%s,String> %s;\n", parameter, name);
    }

    private static String declareAccessor(String fieldName)
    {
        return declareSet(fieldName) + "" + declareGet(fieldName);
    }

    private static String declareSet(String fieldName)
    {
        String name = generateName(fieldName);
        return String.format("    public void set%s(String arg)\n"
                + "    {\n"
                + "        %s = arg;\n"
                + "    }\n", name, fieldName);
    }

    private static String declareGet(String fieldName)
    {
        String name = generateName(fieldName);
        return String.format("    public String get%s()\n"
                + "    {\n"
                + "        return %s;\n"
                + "    }\n", name, fieldName);
    }

    private static String generateName(String field)
    {
        return Character.toUpperCase(field.charAt(0)) + field.substring(1);
    }

    private static String generateGetName(String field)
    {
        return "get" + Character.toUpperCase(field.charAt(0)) + field.substring(1);
    }

    private static String getCellValue(XSSFCell cell)
    {
        switch (cell.getCellTypeEnum())
        {
            case STRING:
                return cell.getStringCellValue().trim();
            case BOOLEAN:
                return cell.getBooleanCellValue() + "";
            case ERROR:
                return cell.getErrorCellString() + "";
            case NUMERIC:
                return cell.getNumericCellValue() + "";
            case BLANK:
                return "";
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

}
