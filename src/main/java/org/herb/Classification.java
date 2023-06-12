package org.herb;

import org.apache.commons.collections4.Bag;
import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.bag.HashBag;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.commons.math3.random.RandomDataGenerator;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.commons.math3.util.Precision;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartFrame;
import org.jfree.chart.JFreeChart;
import org.jfree.data.statistics.BoxAndWhiskerCategoryDataset;
import org.jfree.data.statistics.DefaultBoxAndWhiskerCategoryDataset;
import weka.classifiers.Evaluation;
import weka.classifiers.trees.J48;
import weka.classifiers.trees.RandomForest;
import weka.core.Instance;
import weka.core.Instances;
import weka.core.converters.CSVLoader;
import weka.filters.Filter;
import weka.filters.unsupervised.attribute.NumericToNominal;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Classification {
    public static void main(String[] args) throws Exception {
        String fileLocation = "C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\default of credit card clients.xls";
        String outputFilePath = "C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\output\\default of credit card clients.xls";
        ///////////////////////////////////////////FILE READING
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new HSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);

        HashMap<Integer, List<String>> data = readExcel(sheet);

        head(data);
        System.out.println("Out 2");
        //////////////////////////////////////////PREPROCESSING

        sumOfNullValues(sheet);
        System.out.println("Out 4");
        ///////////////////////////////////////////////////////

        valueCounts(sheet, "SEX");
        System.out.println("Out 5");
        ///////////////////////////////////////////////////////

        Map<String, Double> valueMappings = new HashMap<>();
        valueMappings.put("female", 0.0);
        valueMappings.put("male", 1.0);

        // Map the values in the column
        mapColumnValues(sheet, "SEX", valueMappings);
        // Save the modified workbook to a new file
        FileOutputStream fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);

        System.out.println("Mapping completed successfully.");
        System.out.println("Out 6");
        ///////////////////////////////////////////////////////

        valueCounts(sheet, "SEX");
        System.out.println("Out 7");
        ///////////////////////////////////////////////////////

        Map<String, List<String>> variableMappings = new HashMap<>();
        variableMappings.put("EDUCATION", List.of("dm")); // Specify the column name and the prefix for dummy variables

        // Convert categorical variables into dummy variables
        convertToDummyVariables(sheet, variableMappings);

        // Save the modified workbook to a new file
        fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);


        System.out.println("Dummy variable conversion completed successfully.");
        System.out.println("Out 8");
        ///////////////////////////////////////////////////////

        head(readExcel(sheet));
        System.out.println("Out 9");
        ///////////////////////////////////////////////////////

        valueCounts(sheet, "MARRIAGE");
        System.out.println("Out 10");
        ///////////////////////////////////////////////////////

        int columnIndex = getColumnIndex(sheet, "MARRIAGE");
        if (columnIndex != -1) {
            // Create a new column header
            Row headerRow = sheet.getRow(0);
            Cell newColumnHeader = headerRow.createCell(sheet.getRow(0).getLastCellNum());
            newColumnHeader.setCellValue("MARRIAGE_LE");

            // Perform label encoding
            LabelEncoder labelEncoder = new LabelEncoder();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(columnIndex);
                String value = getCellValueAsString(cell);

                int encodedValue = labelEncoder.fitTransform(value);
                Cell newCell = row.createCell(sheet.getRow(rowIndex).getLastCellNum());
                newCell.setCellValue(encodedValue);
            }

            // Save the modified workbook
            fos = new FileOutputStream(outputFilePath);
            workbook.write(fos);
            System.out.println("Label encoding completed successfully.");
        } else {
            System.out.println("Column name not found in the sheet.");
        }
        head(readExcel(sheet));
        System.out.println("Out 11");
        ///////////////////////////////////////////////////////

        valueCounts(sheet, "MARRIAGE_LE");
        System.out.println("Out 12");
        ///////////////////////////////////////////////////////

        System.out.println(readExcel(sheet).get(0));
        System.out.println("Out 13");
        ///////////////////////////////////////////////////////

        // Define the columns to select and their order
        List<String> columnsToSelect = Arrays.asList(
                "LIMIT_BAL", "SEX", "AGE", "PAY_0", "PAY_2", "PAY_3",
                "PAY_4", "PAY_5", "PAY_6", "BILL_AMT1", "BILL_AMT2", "BILL_AMT3",
                "BILL_AMT4", "BILL_AMT5", "BILL_AMT6", "PAY_AMT1", "PAY_AMT2",
                "PAY_AMT3", "PAY_AMT4", "PAY_AMT5", "PAY_AMT6",
                "dm_graduate school", "dm_high school", "dm_not educated", "dm_others",
                "dm_university", "MARRIAGE_LE", "default"
        );

        // Create a new workbook to store the selected columns
        Workbook newWorkbook = new HSSFWorkbook();
        Sheet newSheet = newWorkbook.createSheet("Selected Columns");

        // Copy the column headers to the new sheet
        Row headerRow = sheet.getRow(0);
        Row newHeaderRow = newSheet.createRow(0);
        for (int i = 0; i < columnsToSelect.size(); i++) {
            String columnName = columnsToSelect.get(i);
            Cell headerCell = headerRow.getCell(getColumnIndex(sheet, columnName));
            Cell newHeaderCell = newHeaderRow.createCell(i);
            newHeaderCell.setCellValue(headerCell.getStringCellValue());
        }

        // Copy the data from the selected columns to the new sheet
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Row newRow = newSheet.createRow(rowIndex);
            for (int i = 0; i < columnsToSelect.size(); i++) {
                String columnName = columnsToSelect.get(i);
                Cell cell = row.getCell(getColumnIndex(sheet, columnName));
                Cell newCell = newRow.createCell(i);
                copyCellValue(cell, newCell);
            }
        }

        // Auto-size the columns in the new sheet
        for (int i = 0; i < columnsToSelect.size(); i++) {
            newSheet.autoSizeColumn(i);
        }

        // Write the new workbook to the output file
        fos = new FileOutputStream(outputFilePath);
        newWorkbook.write(fos);
        System.out.println("Selected columns copied successfully.");
        head(readExcel(newSheet));
        System.out.println("Out 14");
        ///////////////////////////////////////////////////////

        int totalColumns = newSheet.getRow(0).getLastCellNum();
        int lastColumnIndex = totalColumns - 8; // Exclude the last 7 columns

        // Create a list to store the column data
        List<List<Double>> columnDataList = new ArrayList<>();

        // Initialize the list with empty lists
        for (int i = 0; i < lastColumnIndex; i++) {
            columnDataList.add(new ArrayList<>());
        }

        // Extract the column data from the sheet
        for (int rowIndex = 1; rowIndex <= newSheet.getLastRowNum(); rowIndex++) {
            Row row = newSheet.getRow(rowIndex);
            for (int i = 0; i < lastColumnIndex; i++) {
                Cell cell = row.getCell(i);
                double cellValue;

                // Check the cell type and retrieve the cell value accordingly
                if (cell.getCellType() == CellType.NUMERIC) {
                    cellValue = cell.getNumericCellValue();
                } else if (cell.getCellType() == CellType.STRING) {
                    try {
                        cellValue = Double.parseDouble(cell.getStringCellValue());
                    } catch (NumberFormatException e) {
                        cellValue = Double.NaN; // Set NaN for non-numeric values
                    }
                } else {
                    cellValue = Double.NaN; // Set NaN for other cell types
                }

                columnDataList.get(i).add(cellValue);
            }
        }

        // Create a custom BoxAndWhiskerCategoryDataset
        BoxAndWhiskerCategoryDataset dataset = createBoxAndWhiskerDataset(columnDataList, newSheet, lastColumnIndex);

        // Create the box plot chart
        JFreeChart chart = ChartFactory.createBoxAndWhiskerChart(
                "Box Plot",
                "Category",
                "Value",
                dataset,
                true
        );

        // Display the chart in a frame
        ChartFrame frame = new ChartFrame("Box Plot", chart);
        frame.pack();
        frame.setVisible(true);

        System.out.println("Out 15");
        ///////////////////////////////////////////////////////

        int total = newSheet.getRow(0).getLastCellNum();
        int lastCI = total - 7; // Exclude the last 7 columns

        // Calculate quartiles and IQR for each column
        for (int i = 0; i < lastCI; i++) {
            String columnTitle = newSheet.getRow(0).getCell(i).getStringCellValue();

            // Extract column data
            double[] columnData = extractColumnData(sheet, i);

            // Calculate quartiles
            double q1 = calculateQuartile(columnData, 0.25);
            double q3 = calculateQuartile(columnData, 0.75);

            // Calculate IQR
            double iqr = q3 - q1;

            // Calculate lower and upper bounds
            double lowerBound = q1 - 1.5 * iqr;
            double upperBound = q3 + 1.5 * iqr;

            // Treat outliers
            treatOutliers(newSheet, i, lowerBound, upperBound);

            // Plot box plot
            plotBoxPlot(columnData, columnTitle);
        }

        // Save the modified workbook to the output file
        fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);
        System.out.println("Outlier treatment completed. Modified data saved");
        System.out.println("Out 16");
        ///////////////////////////////////////////////////////


        System.out.println(readExcel(newSheet).get(0));
        System.out.println("Out 17");
        ///////////////////////////////////////////////////VIF

        String inputFile = "C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\default of credit card clients 2.xls";

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook wb = new HSSFWorkbook(fis)) {

            sheet = wb.getSheetAt(0);

            List<String> featureNames = List.of("LIMIT_BAL", "SEX", "PAY_0", "PAY_2", "PAY_3", "PAY_4",
                    "PAY_6", "BILL_AMT1", "PAY_AMT1", "PAY_AMT2", "PAY_AMT3", "PAY_AMT4", "PAY_AMT5",
                    "PAY_AMT6", "dm_high school", "dm_not educated", "dm_others", "dm_university",
                    "MARRIAGE_LE");

            List<Double> vifValues = calculateVIF(sheet, featureNames);

            // Print the VIF values and feature names
            System.out.println("VIF\t\tFeatures");
            for (int i = 0; i < featureNames.size(); i++) {
                System.out.println(vifValues.get(i) + "\t\t" + featureNames.get(i));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Out 18");
        ///////////////////////////////////////////////////////

        sheet = workbook.getSheetAt(0);

        List<String> selectedColumns = Arrays.asList("LIMIT_BAL", "SEX", "PAY_0", "PAY_2", "PAY_3",
                "PAY_4", "PAY_6", "BILL_AMT1", "PAY_AMT1", "PAY_AMT2", "PAY_AMT3", "PAY_AMT4",
                "PAY_AMT5", "PAY_AMT6", "dm_high school", "dm_not educated", "dm_others",
                "dm_university", "MARRIAGE_LE", "default");

        Sheet selectedSheet = selectColumns(sheet, selectedColumns, "newSheet");

        head(readExcel(selectedSheet));
        fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);
        System.out.println("Out 19");
        ///////////////////////////////////////////////////////

        System.out.println(readExcel(selectedSheet).get(0));
        System.out.println("Out 20");
        ///////////////////////////////////////////////////////

        double[][] doubles = readExcel("C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\output\\default of credit card clients.xls");

        double[][] correlationMatrix = calculateCorrelationMatrix(doubles);

        printMatrix(correlationMatrix);
        System.out.println("Out 21");
        ///////////////////////////////////////////////////////

        sheet = workbook.getSheetAt(0);

        selectedColumns = Arrays.asList("LIMIT_BAL", "PAY_0", "PAY_2", "PAY_3", "PAY_4", "PAY_6", "default");

        selectedSheet = selectColumns(sheet, selectedColumns, "correlatedSheet");

        doubles = readData(selectedSheet);

        correlationMatrix = calculateCorrelationMatrix(doubles);

        printMatrix(correlationMatrix);

        fos = new FileOutputStream("C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\output\\default of credit card clients 3.xls");
        workbook.write(fos);
        System.out.println("Out 22");
        ///////////////////////////////////////////////////////

        inputFile = "C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\default of credit card clients(correlated).xls";
        FileInputStream fis = new FileInputStream(inputFile);
        workbook = new HSSFWorkbook(fis);

        sheet = workbook.getSheet("correlatedSheet");

        applyHeatmapColors(sheet, correlationMatrix);
        fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);

        System.out.println("Heatmap created successfully!");

        System.out.println("Out 23");
        ///////////////////////////////////////////////////////

        head(readExcel(selectedSheet));
        System.out.println("Out 24");
        ///////////////////////////////////////////////////////

        String excelFilePath = "C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\default of credit card clients 3.xls";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        workbook = new HSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheet("correlatedSheet");

        List<double[]> dataList = readDataFromSheet(firstSheet);

        double[][] array = convertListToArray(dataList);

        double[][] scaledData = scaleData(array);

        double[][] xScaled = extractFeatures(scaledData);
        double[][] y = extractLabels(array);

        double testSize = 0.2;
        long randomSeed = 42;

        int numRows = xScaled.length;
        int numCols = xScaled[0].length;

        int trainSize = (int) Math.round((1.0 - testSize) * numRows);

        RandomDataGenerator randomDataGenerator = new RandomDataGenerator();
        randomDataGenerator.reSeed(randomSeed);

        int[] indices = randomDataGenerator.nextPermutation(numRows, numRows);

        double[][] xTrain = new double[trainSize][numCols];
        double[][] xTest = new double[numRows - trainSize][numCols];
        double[][] yTrain = new double[trainSize][1];
        double[][] yTest = new double[numRows - trainSize][1];

        for (int i = 0; i < numRows; i++) {
            if (i < trainSize) {
                xTrain[i] = xScaled[indices[i]];
                yTrain[i][0] = y[indices[i]][0];
            } else {
                xTest[i - trainSize] = xScaled[indices[i]];
                yTest[i - trainSize][0] = y[indices[i]][0];
            }
        }

        System.out.println("xTrain:");
        printHeadData(xTrain);
        System.out.println("xTest:");
        printHeadData(xTest);
        System.out.println("yTrain:");
        printHeadData(yTrain);
        System.out.println("yTest:");
        printHeadData(yTest);

        System.out.println("Out 25");
        ////////////////////////////////////////Decision Tree Classifier

        // Load the CSV file using Weka's CSVLoader
        CSVLoader loader = new CSVLoader();
        loader.setSource(new File("C:\\Users\\Rafiq\\IdeaProjects\\python_project\\src\\main\\resources\\default of credit card clients 4.csv"));
        Instances instances = loader.getDataSet();

        // Set the target variable (class) index
        instances.setClassIndex(instances.numAttributes() - 1);

        // Convert the class attribute to nominal
        NumericToNominal convert = new NumericToNominal();
        convert.setAttributeIndices(String.valueOf(instances.classIndex() + 1));
        convert.setInputFormat(instances);
        instances = Filter.useFilter(instances, convert);

        // Split the data into train and test subsets
         testSize = 0.2;
         randomSeed = 42;
        Instances trainData = instances.trainCV((int) (1 / testSize), 0, new java.util.Random(randomSeed));
        Instances testData = instances.testCV((int) (1 / testSize), 0);

        // Create and build the decision tree classifier
        J48 classifier = new J48();
        classifier.buildClassifier(trainData);

        // Make predictions on the test data
        int numInstances = testData.numInstances();
        for (int i = 0; i < numInstances; i++) {
            Instance instance = testData.instance(i);
            double predictedClass = classifier.classifyInstance(instance);
        }

        // Evaluate the classifier's performance (e.g., accuracy)
        Evaluation evaluation = new Evaluation(trainData);
        evaluation.evaluateModel(classifier, testData);
        System.out.println("Accuracy: " + evaluation.pctCorrect() + "%");

        // Compute confusion matrix
        int[][] confusionMatrix = computeConfusionMatrix(classifier, testData);
        System.out.println("Confusion Matrix:");
        printMatrix(confusionMatrix);

        // Compute accuracy
        double accuracy = computeAccuracy(confusionMatrix);
        System.out.println("Accuracy: " + accuracy);

        // Compute ROC AUC
        double rocAuc = computeRocAuc(classifier, testData);
        System.out.println("ROC AUC: " + rocAuc);

        // Compute Gini
        double gini = computeGini(rocAuc);
        System.out.println("Gini: " + gini);

        // Compute F1 score
        double f1Score = computeF1Score(classifier, testData);
        System.out.println("F1 Score: " + f1Score);

        System.out.println("Out 26");
        //////////////////////////////////////////Random Forest Classifier

        // Create and build the random forest classifier
        RandomForest randomForest = new RandomForest();
        randomForest.buildClassifier(trainData);

        // Make predictions on the test data
        int nummedInstances = testData.numInstances();
        for (int i = 0; i < nummedInstances; i++) {
            Instance instance = testData.instance(i);
            double predictedClass = randomForest.classifyInstance(instance);
            System.out.println("Predicted class for instance " + (i + 1) + ": " + predictedClass);
        }

        // Evaluate the classifier's performance (e.g., accuracy)
        evaluation = new Evaluation(trainData);
        evaluation.evaluateModel(randomForest, testData);
        System.out.println("Accuracy: " + evaluation.pctCorrect() + "%");

        // Compute confusion matrix
        confusionMatrix = computeConfusionMatrix(randomForest, testData);
        System.out.println("Confusion Matrix:");
        printMatrix(confusionMatrix);

        // Compute accuracy
        accuracy = computeAccuracy(confusionMatrix);
        System.out.println("Accuracy: " + accuracy);

        // Compute ROC AUC
        rocAuc = computeRocAuc(randomForest, testData);
        System.out.println("ROC AUC: " + rocAuc);

        // Compute Gini
        gini = computeGini(rocAuc);
        System.out.println("Gini: " + gini);

        // Compute F1 score
        f1Score = computeF1Score(randomForest, testData);
        System.out.println("F1 Score: " + f1Score);

        System.out.println("Out 27");
    }

    private static int[][] computeConfusionMatrix(RandomForest classifier, Instances testData) throws Exception {
        int[][] confusionMatrix = new int[testData.numClasses()][testData.numClasses()];

        for (int i = 0; i < testData.numInstances(); i++) {
            Instance instance = testData.instance(i);
            double trueLabel = instance.classValue();
            double predictedLabel = classifier.classifyInstance(instance);

            confusionMatrix[(int) trueLabel][(int) predictedLabel]++;
        }

        return confusionMatrix;
    }


    private static double computeRocAuc(RandomForest classifier, Instances testData) throws Exception {
        Evaluation evaluation = new Evaluation(testData);
        evaluation.evaluateModel(classifier, testData);

        return evaluation.areaUnderROC(1);
    }
    private static double computeF1Score(RandomForest classifier, Instances testData) throws Exception {
        int truePositives = 0;
        int falsePositives = 0;
        int falseNegatives = 0;

        for (int i = 0; i < testData.numInstances(); i++) {
            Instance instance = testData.instance(i);
            double trueLabel = instance.classValue();
            double predictedLabel = classifier.classifyInstance(instance);

            if (trueLabel == 1 && predictedLabel == 1) {
                truePositives++;
            } else if (trueLabel == 0 && predictedLabel == 1) {
                falsePositives++;
            } else if (trueLabel == 1 && predictedLabel == 0) {
                falseNegatives++;
            }
        }

        double precision = (double) truePositives / (truePositives + falsePositives);
        double recall = (double) truePositives / (truePositives + falseNegatives);
        return 2 * precision * recall / (precision + recall);
    }
    private static int[][] computeConfusionMatrix(J48 classifier, Instances testData) throws Exception {
        int[][] confusionMatrix = new int[testData.numClasses()][testData.numClasses()];

        for (int i = 0; i < testData.numInstances(); i++) {
            Instance instance = testData.instance(i);
            double trueLabel = instance.classValue();
            double predictedLabel = classifier.classifyInstance(instance);

            confusionMatrix[(int) trueLabel][(int) predictedLabel]++;
        }

        return confusionMatrix;
    }

    private static double computeAccuracy(int[][] confusionMatrix) {
        int correct = 0;
        int total = 0;

        for (int i = 0; i < confusionMatrix.length; i++) {
            for (int j = 0; j < confusionMatrix[i].length; j++) {
                if (i == j) {
                    correct += confusionMatrix[i][j];
                }
                total += confusionMatrix[i][j];
            }
        }

        return (double) correct / total;
    }

    private static double computeRocAuc(J48 classifier, Instances testData) throws Exception {
        Evaluation evaluation = new Evaluation(testData);
        evaluation.evaluateModel(classifier, testData);

        return evaluation.areaUnderROC(1);
    }

    private static double computeGini(double rocAuc) {
        return (2 * rocAuc) - 1;
    }

    private static double computeF1Score(J48 classifier, Instances testData) throws Exception {
        int truePositives = 0;
        int falsePositives = 0;
        int falseNegatives = 0;

        for (int i = 0; i < testData.numInstances(); i++) {
            Instance instance = testData.instance(i);
            double trueLabel = instance.classValue();
            double predictedLabel = classifier.classifyInstance(instance);

            if (trueLabel == 1 && predictedLabel == 1) {
                truePositives++;
            } else if (trueLabel == 0 && predictedLabel == 1) {
                falsePositives++;
            } else if (trueLabel == 1 && predictedLabel == 0) {
                falseNegatives++;
            }
        }

        double precision = (double) truePositives / (truePositives + falsePositives);
        double recall = (double) truePositives / (truePositives + falseNegatives);
        return 2 * precision * recall / (precision + recall);
    }

    private static void printMatrix(int[][] matrix) {
        for (int i = 0; i < matrix.length; i++) {
            for (int j = 0; j < matrix[i].length; j++) {
                System.out.print(matrix[i][j] + " ");
            }
            System.out.println();
        }
    }

    private static List<double[]> readDataFromSheet(Sheet sheet) {
        List<double[]> dataList = new ArrayList<>();

        for (Row row : sheet) {
            double[] rowData = new double[row.getLastCellNum()];
            for (int i = 0; i < row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i);
                rowData[i] = cell.getNumericCellValue();
            }
            dataList.add(rowData);
        }

        return dataList;
    }

    private static double[][] convertListToArray(List<double[]> dataList) {
        int numRows = dataList.size();
        int numCols = dataList.get(0).length;

        double[][] data = new double[numRows][numCols];

        for (int i = 0; i < numRows; i++) {
            data[i] = dataList.get(i);
        }

        return data;
    }

    private static double[][] scaleData(double[][] data) {
        int numRows = data.length;
        int numCols = data[0].length;

        double[][] scaledData = new double[numRows][numCols];

        double[] means = calculateMeans(data);
        double[] stdDevs = calculateStandardDeviations(data, means);

        for (int i = 0; i < numRows; i++) {
            for (int j = 0; j < numCols; j++) {
                scaledData[i][j] = (data[i][j] - means[j]) / stdDevs[j];
            }
        }

        return scaledData;
    }

    private static double[] calculateMeans(double[][] data) {
        int numCols = data[0].length;
        double[] means = new double[numCols];

        for (int j = 0; j < numCols; j++) {
            double sum = 0.0;
            for (double[] row : data) {
                sum += row[j];
            }
            means[j] = sum / data.length;
        }

        return means;
    }

    private static double[] calculateStandardDeviations(double[][] data, double[] means) {
        int numCols = data[0].length;
        double[] stdDevs = new double[numCols];

        for (int j = 0; j < numCols; j++) {
            double sumSquaredDiff = 0.0;
            for (double[] row : data) {
                double diff = row[j] - means[j];
                sumSquaredDiff += diff * diff;
            }
            stdDevs[j] = Math.sqrt(sumSquaredDiff / data.length);
        }

        return stdDevs;
    }

    private static double[][] extractFeatures(double[][] data) {
        int numRows = data.length;
        int numCols = data[0].length - 1; // Exclude the last column for labels

        double[][] features = new double[numRows][numCols];

        for (int i = 0; i < numRows; i++) {
            for (int j = 0; j < numCols; j++) {
                features[i][j] = data[i][j];
            }
        }

        return features;
    }

    private static double[][] extractLabels(double[][] data) {
        int numRows = data.length;
        int numCols = 1; // Only one column for labels

        double[][] labels = new double[numRows][numCols];

        for (int i = 0; i < numRows; i++) {
            labels[i][0] = data[i][data[0].length - 1]; // Last column for labels
        }

        return labels;
    }

    private static void printHeadData(double[][] data) {
        for (int i = 0; i < 6; i++) {
            double[] row = data[i];
            for (int j = 0; j < row.length; j++) {
                double value = row[j];
                System.out.print(Precision.round(value, 6) + "\t");
            }
            System.out.println();
        }
    }

    private static void applyHeatmapColors(Sheet sheet, double[][] correlationMatrix) {
        int numRows = correlationMatrix.length;
        int numCols = correlationMatrix[0].length;

        // Set the background color for each cell based on the correlation value
        for (int row = 0; row < numRows; row++) {
            Row currentRow = sheet.getRow(row);
            if (currentRow == null) {
                currentRow = sheet.createRow(row);
            }

            for (int col = 0; col < numCols; col++) {
                double correlation = correlationMatrix[row][col];
                Cell cell = currentRow.getCell(col);
                if (cell == null) {
                    cell = currentRow.createCell(col);
                }

                // Set the color based on the correlation value
                if (correlation >= 0.7) {
                    cell.setCellStyle(getCellStyle(sheet, IndexedColors.RED));
                } else if (correlation >= 0.4) {
                    cell.setCellStyle(getCellStyle(sheet, IndexedColors.YELLOW));
                } else if (correlation >= -0.4) {
                    cell.setCellStyle(getCellStyle(sheet, IndexedColors.GREEN));
                } else {
                    cell.setCellStyle(getCellStyle(sheet, IndexedColors.BLUE));
                }
            }
        }
    }

    private static CellStyle getCellStyle(Sheet sheet, IndexedColors color) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(color.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    private static double[][] calculateCorrelationMatrix(double[][] data) {
        int numRows = data.length;
        int numCols = data[0].length;

        double[][] correlationMatrix = new double[numCols][numCols];

        for (int col1 = 0; col1 < numCols; col1++) {
            for (int col2 = 0; col2 < numCols; col2++) {
                double sumX = 0.0;
                double sumY = 0.0;
                double sumXY = 0.0;
                double sumX2 = 0.0;
                double sumY2 = 0.0;
                int count = 0;

                for (int row = 0; row < numRows; row++) {
                    double x = data[row][col1];
                    double y = data[row][col2];

                    if (!Double.isNaN(x) && !Double.isNaN(y)) {
                        sumX += x;
                        sumY += y;
                        sumXY += x * y;
                        sumX2 += x * x;
                        sumY2 += y * y;
                        count++;
                    }
                }

                double correlation = calculateCorrelation(sumX, sumY, sumXY, sumX2, sumY2, count);
                correlationMatrix[col1][col2] = correlation;
            }
        }

        return correlationMatrix;
    }

    private static double calculateCorrelation(double sumX, double sumY, double sumXY, double sumX2, double sumY2,
                                               int count) {
        double numerator = (count * sumXY) - (sumX * sumY);
        double denominator = Math.sqrt((count * sumX2 - Math.pow(sumX, 2)) * (count * sumY2 - Math.pow(sumY, 2)));

        if (denominator != 0.0) {
            return numerator / denominator;
        } else {
            return 0.0;
        }
    }

    private static void printMatrix(double[][] matrix) {
        int numRows = matrix.length;
        int numCols = matrix[0].length;

        for (int row = 0; row < numRows; row++) {
            for (int col = 0; col < numCols; col++) {
                System.out.print(matrix[row][col] + "\t");
            }
            System.out.println();
        }
    }

    private static Sheet selectColumns(Sheet sheet, List<String> selectedColumns, String newSheetName) {
        Workbook workbook = sheet.getWorkbook();
        Sheet selectedSheet = workbook.createSheet(newSheetName);

        // Copy the header row to the selected sheet
        Row headerRow = sheet.getRow(0);
        Row selectedHeaderRow = selectedSheet.createRow(0);
        for (int columnIndex = 0; columnIndex < headerRow.getLastCellNum(); columnIndex++) {
            Cell headerCell = headerRow.getCell(columnIndex);
            String columnName = headerCell.getStringCellValue();
            if (selectedColumns.contains(columnName)) {
                Cell selectedHeaderCell = selectedHeaderRow.createCell(selectedColumns.indexOf(columnName));
                selectedHeaderCell.setCellValue(columnName);
            }
        }

        // Copy the data rows to the selected sheet
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Row selectedRow = selectedSheet.createRow(rowIndex);
            for (int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    String columnName = headerRow.getCell(columnIndex).getStringCellValue();
                    if (selectedColumns.contains(columnName)) {
                        Cell selectedCell = selectedRow.createCell(selectedColumns.indexOf(columnName));
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case NUMERIC:
                                selectedCell.setCellValue(cell.getNumericCellValue());
                                break;
                            case STRING:
                                selectedCell.setCellValue(cell.getStringCellValue());
                                break;
                            case BOOLEAN:
                                selectedCell.setCellValue(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                selectedCell.setCellFormula(cell.getCellFormula());
                                break;
                            // Handle other cell types if needed
                        }
                    }
                }
            }
        }

        return selectedSheet;
    }

    private static List<Double> calculateVIF(Sheet sheet, List<String> featureNames) {
        List<Double> vifValues = new ArrayList<>();

        int numRows = sheet.getLastRowNum() + 1;
        int numFeatures = featureNames.size();

        // Read feature values from the Excel sheet
        double[][] features = new double[numRows][numFeatures];
        for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            for (int columnIndex = 0; columnIndex < numFeatures; columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    features[rowIndex][columnIndex] = cell.getNumericCellValue();
                }
            }
        }

        // Calculate VIF for each feature
        for (int i = 0; i < numFeatures; i++) {
            double[] x = new double[numRows];
            for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
                x[rowIndex] = features[rowIndex][i];
            }
            double vif = calculateVIF(x);
            vifValues.add(vif);
        }

        return vifValues;
    }

    private static double calculateVIF(double[] x) {
        int n = x.length;
        double[] residuals = new double[n];
        double[] predicted = new double[n];

        double meanX = calculateMean(x);

        // Calculate predicted values using simple linear regression
        for (int i = 0; i < n; i++) {
            predicted[i] = calculateLinearRegression(x, i, meanX);
        }

        // Calculate residuals
        for (int i = 0; i < n; i++) {
            residuals[i] = x[i] - predicted[i];
        }

        double sse = calculateSumOfSquares(residuals);
        double mse = sse / (n - 2); // Mean squared error
        double varX = calculateVariance(x);

        // Calculate VIF
        double vif = varX / mse;

        return vif;
    }

    private static double calculateMean(double[] values) {
        double sum = 0.0;
        for (double value : values) {
            sum += value;
        }
        return sum / values.length;
    }

    private static double calculateLinearRegression(double[] x, int index, double meanX) {
        double sumXY = 0.0;
        double sumX2 = 0.0;
        int n = x.length;

        for (int i = 0; i < n; i++) {
            if (i != index) {
                sumXY += (x[i] - meanX) * (x[index] - meanX);
                sumX2 += Math.pow((x[i] - meanX), 2);
            }
        }

        return sumXY / sumX2 * x[index];
    }

    private static double calculateSumOfSquares(double[] values) {
        double sum = 0.0;
        for (double value : values) {
            sum += Math.pow(value, 2);
        }
        return sum;
    }

    private static double calculateVariance(double[] values) {
        double mean = calculateMean(values);
        double sum = 0.0;
        int n = values.length;

        for (double value : values) {
            sum += Math.pow((value - mean), 2);
        }

        return sum / (n - 1);
    }

    public static double[][] readExcel(String filePath) {
        try {
            FileInputStream excelFile = new FileInputStream(filePath);
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            int numOfRows = datatypeSheet.getPhysicalNumberOfRows();
            int numOfCols = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
            double[][] data = new double[numOfRows][numOfCols];

            int rowNum = 0;
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                int colNum = 0;
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellType() == CellType.NUMERIC) {
                        data[rowNum][colNum] = currentCell.getNumericCellValue();
                    }
                    colNum++;
                }
                rowNum++;
            }
            return data;

        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private static double[][] readData(Sheet sheet) {
        int numRows = sheet.getLastRowNum() + 1;
        int numCols = sheet.getRow(0).getLastCellNum();

        double[][] data = new double[numRows][numCols];

        for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            for (int colIndex = 0; colIndex < numCols; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    data[rowIndex][colIndex] = cell.getNumericCellValue();
                }
            }
        }

        return data;
    }

    public static HashMap<Integer, List<String>> readExcel(Sheet sheet) {
        HashMap<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING:
                        data.get(i).add(cell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            data.get(i).add(String.valueOf(cell.getDateCellValue()));
                        } else {
                            data.get(i).add(String.valueOf(cell.getNumericCellValue()));
                        }
                        break;
                    case BOOLEAN:
                        data.get(i).add(String.valueOf(cell.getBooleanCellValue()));
                        break;
                    case FORMULA:
                        data.get(i).add(String.valueOf(cell.getCellFormula()));
                        break;
                    default:
                        data.get(i).add(" ");
                }
            }
            i++;
        }
        return data;
    }

    public static void head(HashMap<Integer, List<String>> data) {
        for (int i = 0; i < 6; i++) {
            System.out.println(data.get(i));
        }
    }

    public static void sumOfNullValues(Sheet sheet) {
        Row headerRow = sheet.getRow(0); // Get the first row as header row

        int totalColumns = headerRow.getLastCellNum(); // Assuming the first row contains column headers

        // Initialize an array to keep track of null values count for each column
        int[] columnNullCount = new int[totalColumns];

        // Iterate over each row in the sheet
        for (Row row : sheet) {
            // Iterate over each cell in the row
            for (int columnIndex = 0; columnIndex < totalColumns; columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    // Null or blank cell found
                    columnNullCount[columnIndex]++;
                }
            }
        }

        // Print the column name and count of null values for each column
        for (int columnIndex = 0; columnIndex < totalColumns; columnIndex++) {
            Cell headerCell = headerRow.getCell(columnIndex);
            String columnName = headerCell.getStringCellValue();
            int nullCount = columnNullCount[columnIndex];
            System.out.println("Column " + columnName + " has " + nullCount + " null values.");
        }
    }

    public static void valueCounts(Sheet sheet, String columnName) {
        Bag<String> rowCounts = new HashBag<>();

        // Iterate over each row in the sheet
        for (Row row : sheet) {
            // Skip the header row
            if (row.getRowNum() == 0) {
                continue;
            }

            int columnIndex = getColumnIndex(sheet, columnName);
            // Get the cell value from the specified column
            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValueAsString(cell);

            // Count row occurrences
            rowCounts.add(cellValue);
        }

        // Print counts of unique rows
        for (String rowValue : rowCounts.uniqueSet()) {
            int count = rowCounts.getCount(rowValue);
            System.out.println("Row value: " + rowValue + ", Count: " + count);
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else {
            return "";
        }
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0); // Assuming the first row contains column headers
        int lastCellNum = headerRow.getLastCellNum();
        for (int columnIndex = 0; columnIndex < lastCellNum; columnIndex++) {
            Cell cell = headerRow.getCell(columnIndex);
            String headerValue = getCellValueAsString(cell);
            if (headerValue.equalsIgnoreCase(columnName)) {
                return columnIndex;
            }
        }
        return -1; // Column name not found
    }

    public static void mapColumnValues(Sheet sheet, String columnName, Map<String, Double> valueMappings) {
        for (Row row : sheet) {
            // Skip the header row
            if (row.getRowNum() == 0) {
                continue;
            }
            int columnIndex = getColumnIndex(sheet, columnName);
            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValueAsString(cell);

            // Map the cell value using the valueMappings map
            if (valueMappings.containsKey(cellValue)) {
                double mappedValue = valueMappings.get(cellValue);
                cell.setCellValue(mappedValue);
            }
        }
    }

    private static void convertToDummyVariables(Sheet sheet, Map<String, List<String>> variableMappings) {
        for (Map.Entry<String, List<String>> entry : variableMappings.entrySet()) {
            String columnName = entry.getKey();
            List<String> prefixes = entry.getValue();

            int columnIndex = getColumnIndex(sheet, columnName);
            if (columnIndex != -1) {
                MultiValuedMap<Integer, String> dummyVariables = new ArrayListValuedHashMap<>();

                // Extract unique values from the column
                Set<String> uniqueValues = getColumnUniqueValues(sheet, columnIndex);

                // Create dummy variable columns with the specified prefixes
                for (String prefix : prefixes) {
                    for (String value : uniqueValues) {
                        String dummyVariableName = prefix + "_" + value;
                        int dummyColumnIndex = createDummyVariableColumn(sheet, dummyVariableName);
                        dummyVariables.put(dummyColumnIndex, value);
                    }
                }

                // Populate the dummy variable columns based on the original column values
                populateDummyVariableColumns(sheet, columnIndex, dummyVariables);
            }
        }
    }

    private static Set<String> getColumnUniqueValues(Sheet sheet, int columnIndex) {
        Set<String> uniqueValues = new HashSet<>();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue; // Skip the header row
            }

            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValueAsString(cell);

            uniqueValues.add(cellValue);
        }

        return uniqueValues;
    }

    private static int createDummyVariableColumn(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        int columnIndex = headerRow.getLastCellNum();
        Cell headerCell = headerRow.createCell(columnIndex);
        headerCell.setCellValue(columnName);
        return columnIndex;
    }

    private static void populateDummyVariableColumns(Sheet sheet, int columnIndex,
                                                     MultiValuedMap<Integer, String> dummyVariables) {
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue; // Skip the header row
            }

            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValueAsString(cell);

            for (Map.Entry<Integer, String> entry : dummyVariables.entries()) {
                int dummyColumnIndex = entry.getKey();
                String value = entry.getValue();

                Cell dummyCell = row.createCell(dummyColumnIndex);
                dummyCell.setCellValue(cellValue.equalsIgnoreCase(value) ? 1 : 0);
            }
        }
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null) {
            return;
        }

        CellType cellType = sourceCell.getCellType();
        switch (cellType) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }

    private static BoxAndWhiskerCategoryDataset createBoxAndWhiskerDataset(List<List<Double>> columnDataList, Sheet sheet, int lastColumnIndex) {
        DefaultBoxAndWhiskerCategoryDataset dataset = new DefaultBoxAndWhiskerCategoryDataset();

        for (int columnIndex = 0; columnIndex < lastColumnIndex; columnIndex++) {
            String columnTitle = sheet.getRow(0).getCell(columnIndex).getStringCellValue();
            List<Double> columnData = columnDataList.get(columnIndex);

            // Compute statistics using Apache Commons Math
            DescriptiveStatistics stats = new DescriptiveStatistics();
            columnData.forEach(stats::addValue);

            // Create a list for the dataset values
            List<Double> datasetValues = new ArrayList<>();
            double[] values = stats.getValues();
            for (double value : values) {
                datasetValues.add(value);
            }

            // Add the dataset values to the custom dataset
            dataset.add(datasetValues, columnTitle, "");

        }
        return dataset;
    }

    private static double[] extractColumnData(Sheet sheet, int columnIndex) {
        int totalRows = sheet.getLastRowNum() + 1;
        double[] columnData = new double[totalRows - 1]; // Exclude the header row

        for (int rowIndex = 1; rowIndex < totalRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);

            // Check the cell type and retrieve the cell value accordingly
            if (cell.getCellType() == CellType.NUMERIC) {
                columnData[rowIndex - 1] = cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                try {
                    columnData[rowIndex - 1] = Double.parseDouble(cell.getStringCellValue());
                } catch (NumberFormatException e) {
                    columnData[rowIndex - 1] = Double.NaN; // Set NaN for non-numeric values
                }
            } else {
                columnData[rowIndex - 1] = Double.NaN; // Set NaN for other cell types
            }
        }

        return columnData;
    }

    private static double calculateQuartile(double[] data, double percentile) {
        double[] sortedData = data.clone();
        Arrays.sort(sortedData);

        double n = (sortedData.length - 1) * percentile + 1;
        int k = (int) n;
        double d = n - k;

        double q = sortedData[k - 1];
        if (k < sortedData.length) {
            q += d * (sortedData[k] - sortedData[k - 1]);
        }

        return q;
    }

    private static void treatOutliers(Sheet sheet, int columnIndex, double lowerBound, double upperBound) {
        int totalRows = sheet.getLastRowNum() + 1;

        for (int rowIndex = 1; rowIndex < totalRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndex);

            double cellValue;

            // Check the cell type and retrieve the cell value accordingly
            if (cell.getCellType() == CellType.NUMERIC) {
                cellValue = cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                try {
                    cellValue = Double.parseDouble(cell.getStringCellValue());
                } catch (NumberFormatException e) {
                    cellValue = Double.NaN; // Set NaN for non-numeric values
                }
            } else {
                cellValue = Double.NaN; // Set NaN for other cell types
            }

            if (cellValue > upperBound) {
                cell.setCellValue(upperBound);
            } else if (cellValue < lowerBound) {
                cell.setCellValue(lowerBound);
            }
        }
    }

    private static void plotBoxPlot(double[] data, String columnTitle) {
        DefaultBoxAndWhiskerCategoryDataset dataset = new DefaultBoxAndWhiskerCategoryDataset();

        // Create a series for the box plot
        List<Double> values = new ArrayList<>();
        for (double value : data) {
            values.add(value);
        }
        dataset.add(values, "Box Plot", columnTitle);

        JFreeChart chart = ChartFactory.createBoxAndWhiskerChart(
                "Box Plot - " + columnTitle,
                "Category",
                "Value",
                dataset,
                true
        );

        ChartFrame frame = new ChartFrame(columnTitle, chart);
        frame.pack();
        frame.setVisible(true);
    }

    private static class LabelEncoder {
        private Map<String, Integer> labelEncoderMap;
        private int encodedValueCounter;

        public LabelEncoder() {
            labelEncoderMap = new HashMap<>();
            encodedValueCounter = 0;
        }

        public int fitTransform(String label) {
            if (labelEncoderMap.containsKey(label)) {
                return labelEncoderMap.get(label);
            } else {
                labelEncoderMap.put(label, encodedValueCounter);
                encodedValueCounter++;
                return encodedValueCounter - 1;
            }
        }
    }
}