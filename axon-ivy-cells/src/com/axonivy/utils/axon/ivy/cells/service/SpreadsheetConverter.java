package com.axonivy.utils.axon.ivy.cells.service;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import ch.ivyteam.ivy.environment.Ivy;

/**
 * Fluent API for spreadsheet conversion operations. Provides a chain of methods to
 * convert spreadsheets from one format to another.
 */
public class SpreadsheetConverter {
  private Workbook workbook;
  private Integer targetFormat;

  /**
   * Creates a new SpreadsheetConverter instance. Package-private constructor to
   * ensure creation only through ExcelFactory.
   */
  SpreadsheetConverter() {
  }

  /**
   * Sets the source spreadsheet from an InputStream.
   * 
   * @param inputStream the input stream containing the spreadsheet data
   * @return this converter instance for method chaining
   * @throws SpreadsheetConversionException if spreadsheet loading fails
   */
  public SpreadsheetConverter from(InputStream inputStream) {
    try {
      this.workbook = new Workbook(inputStream);
      return this;
    } catch (Exception e) {
      Ivy.log().error("Failed to load spreadsheet from InputStream", e);
      throw new SpreadsheetConversionException("Failed to load spreadsheet", e);
    }
  }

  /**
   * Sets the source spreadsheet from a File.
   * 
   * @param file the file containing the spreadsheet
   * @return this converter instance for method chaining
   * @throws SpreadsheetConversionException if spreadsheet loading fails
   */
  public SpreadsheetConverter from(File file) {
    try {
      this.workbook = new Workbook(file.getAbsolutePath());
      return this;
    } catch (Exception e) {
      Ivy.log().error("Failed to load spreadsheet from file: " + file.getAbsolutePath(), e);
      throw new SpreadsheetConversionException("Failed to load spreadsheet from file", e);
    }
  }

  /**
   * Sets the source spreadsheet from a file path.
   * 
   * @param filePath the path to the file containing the spreadsheet
   * @return this converter instance for method chaining
   * @throws SpreadsheetConversionException if spreadsheet loading fails
   */
  public SpreadsheetConverter from(String filePath) {
    try {
      this.workbook = new Workbook(filePath);
      return this;
    } catch (Exception e) {
      Ivy.log().error("Failed to load spreadsheet from path: " + filePath, e);
      throw new SpreadsheetConversionException("Failed to load spreadsheet from path", e);
    }
  }

  /**
   * Sets the source spreadsheet from a byte array.
   * 
   * @param bytes the byte array containing the spreadsheet data
   * @return this converter instance for method chaining
   * @throws SpreadsheetConversionException if spreadsheet loading fails
   */
  public SpreadsheetConverter from(byte[] bytes) {
    try {
      this.workbook = new Workbook(new ByteArrayInputStream(bytes));
      return this;
    } catch (Exception e) {
      Ivy.log().error("Failed to load spreadsheet from byte array", e);
      throw new SpreadsheetConversionException("Failed to load spreadsheet from byte array", e);
    }
  }

  /**
   * Converts the spreadsheet to PDF format.
   * 
   * @return this converter instance for method chaining
   */
  public SpreadsheetConverter toPdf() {
    return to(SaveFormat.PDF);
  }

  /**
   * Converts the spreadsheet to XLSX format.
   * 
   * @return this converter instance for method chaining
   */
  public SpreadsheetConverter toXlsx() {
    return to(SaveFormat.XLSX);
  }

  /**
   * Converts the spreadsheet to XLS format.
   * 
   * @return this converter instance for method chaining
   */
  public SpreadsheetConverter toXls() {
    return to(SaveFormat.EXCEL_97_TO_2003);
  }

  /**
   * Converts the spreadsheet to CSV format.
   * 
   * @return this converter instance for method chaining
   */
  public SpreadsheetConverter toCsv() {
    return to(SaveFormat.CSV);
  }

  /**
   * Converts the spreadsheet to the specified format.
   * 
   * @param format the target format
   * @return this converter instance for method chaining
   */
  public SpreadsheetConverter to(int format) {
    if (workbook == null) {
      throw new IllegalStateException("No source spreadsheet set. Call from() method first.");
    }
    this.targetFormat = format;
    return this;
  }

  /**
   * Converts the spreadsheet and returns the result as a byte array.
   * 
   * @return the converted spreadsheet as byte array
   * @throws SpreadsheetConversionException if conversion fails
   */
  public byte[] asBytes() {
    validateConversionReady();
    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      workbook.save(outputStream, targetFormat);
      return outputStream.toByteArray();
    } catch (Exception e) {
      Ivy.log().error("Failed to convert spreadsheet", e);
      throw new SpreadsheetConversionException("Failed to convert spreadsheet", e);
    }
  }

  /**
   * Converts the spreadsheet and saves it as a file.
   * 
   * @param outputPath the path where the converted file should be saved
   * @return the File object representing the saved file
   * @throws SpreadsheetConversionException if conversion or file saving fails
   */
  public File asFile(String outputPath) {
    validateConversionReady();
    try {
      File outputFile = new File(outputPath);
      // Ensure parent directories exist
      File parentDir = outputFile.getParentFile();
      if (parentDir != null && !parentDir.exists()) {
        parentDir.mkdirs();
      }

      workbook.save(outputPath, targetFormat);
      return outputFile;
    } catch (Exception e) {
      Ivy.log().error("Failed to save converted spreadsheet to: " + outputPath, e);
      throw new SpreadsheetConversionException("Failed to save converted spreadsheet", e);
    }
  }

  /**
   * Converts the spreadsheet and saves it as a file.
   * 
   * @param outputFile the File object where the converted spreadsheet should be
   *                   saved
   * @return the File object representing the saved file
   * @throws SpreadsheetConversionException if conversion or file saving fails
   */
  public File asFile(File outputFile) {
    return asFile(outputFile.getAbsolutePath());
  }

  /**
   * Converts the spreadsheet and returns it as an InputStream. Note: The caller is
   * responsible for closing the returned InputStream.
   * 
   * @return an InputStream containing the converted spreadsheet data
   * @throws SpreadsheetConversionException if conversion fails
   */
  public InputStream asInputStream() {
    byte[] bytes = asBytes();
    return new ByteArrayInputStream(bytes);
  }

  /**
   * Validates that the converter is ready for conversion.
   * 
   * @throws IllegalStateException if workbook or target format is not set
   */
  private void validateConversionReady() {
    if (workbook == null) {
      throw new IllegalStateException("No source spreadsheet set. Call from() method first.");
    }
    if (targetFormat == null) {
      throw new IllegalStateException("No target format set. Call to() or toPdf() method first.");
    }
  }
}
