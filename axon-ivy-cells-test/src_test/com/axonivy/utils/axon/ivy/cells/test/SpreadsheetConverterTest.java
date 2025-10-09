package com.axonivy.utils.axon.ivy.cells.test;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.MockedConstruction;
import org.mockito.Mockito;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.axonivy.utils.axon.ivy.cells.service.SpreadsheetConversionException;
import com.axonivy.utils.axon.ivy.cells.service.SpreadsheetConverter;
import com.axonivy.utils.axon.ivy.cells.service.CellFactory;

import ch.ivyteam.ivy.environment.IvyTest;

@IvyTest
public class SpreadsheetConverterTest {
  private final String TEST_FILE_PATH = "src_test/resources/demo.xlsx";
  private final String TEST_OUTPUT_FILE_PATH = "src_test/resources/output.xlsx";

  @BeforeEach
  void resetLicenseField() throws Exception {
    var field = CellFactory.class.getDeclaredField("license");
    field.setAccessible(true);
    field.set(null, null);
  }

  @Test
  void testConvertFromInputStreamToPdfAsBytes() throws Exception {
    byte[] testData = getDemoWorkbookAsBytes();
    InputStream inputStream = new ByteArrayInputStream(testData);
    byte[] result = CellFactory.convert().from(inputStream).toPdf().asBytes();
    assertNotNull(result);
    assertTrue(result.length > 0);
  }

  @Test
  void testConvertFromFileToFormatAsFile() throws Exception {
    File inputFile = new File(TEST_FILE_PATH);
    File result = CellFactory.convert().from(inputFile).to(SaveFormat.PDF).asFile(TEST_OUTPUT_FILE_PATH);
    assertNotNull(result);
  }

  @Test
  void testConvertFromFilePathToPdfAsInputStream() throws Exception {
    InputStream result = CellFactory.convert().from(TEST_FILE_PATH).toPdf().asInputStream();
    assertNotNull(result);
  }

  @Test
  void testConvertFromBytesArrayToPdfAsBytes() throws Exception {
    byte[] inputBytes = createTestWorkbookAsBytes();
    byte[] result = CellFactory.convert().from(inputBytes).toPdf().asBytes();
    assertNotNull(result);
    assertTrue(result.length > 0);
  }

  /**
   * Creates a test Excel workbook as byte array for testing purposes. This
   * generates a proper Aspose Workbook with sample data.
   */
  private byte[] createTestWorkbookAsBytes() throws Exception {
    Workbook workbook = new Workbook();
    workbook.getWorksheets().get(0).getCells().get("A1").putValue("Test Data");
    workbook.getWorksheets().get(0).getCells().get("A2").putValue(123.45);
    workbook.getWorksheets().get(0).getCells().get("B1").putValue("Sample Text");
    workbook.getWorksheets().get(0).getCells().get("B2").putValue(67.89);

    java.io.ByteArrayOutputStream outputStream = new java.io.ByteArrayOutputStream();
    workbook.save(outputStream, SaveFormat.XLSX);
    return outputStream.toByteArray();
  }

  /**
   * Alternative method to get bytes from existing demo.xlsx file
   */
  private byte[] getDemoWorkbookAsBytes() throws Exception {
    java.nio.file.Path path = java.nio.file.Paths.get(TEST_FILE_PATH);
    return java.nio.file.Files.readAllBytes(path);
  }

  @Test
  void testConvertFromDemoFileBytesToPdf() throws Exception {
    byte[] inputBytes = getDemoWorkbookAsBytes();
    byte[] result = CellFactory.convert().from(inputBytes).toPdf().asBytes();
    assertNotNull(result);
    assertTrue(result.length > 0);
  }

  @Test
  void testConvertWithoutSourceThrowsException() {
    SpreadsheetConverter converter = CellFactory.convert();
    assertThrows(IllegalStateException.class, () -> {
      converter.toPdf().asBytes();
    });
  }

  @Test
  void testConvertFromInvalidInputStreamThrowsException() throws Exception {
    withMockedDocumentFailure(() -> {
      InputStream inputStream = new ByteArrayInputStream("invalid content".getBytes());
      assertThrows(SpreadsheetConversionException.class, () -> {
        CellFactory.convert().from(inputStream).toPdf().asBytes();
      });
    });
  }

  @Test
  void testConvertFromInvalidFileThrowsException() throws Exception {
    withMockedDocumentFailure(() -> {
      File invalidFile = new File("nonexistent.xlsx");
      assertThrows(SpreadsheetConversionException.class, () -> {
        CellFactory.convert().from(invalidFile).toPdf().asBytes();
      });
    });
  }

  private void withMockedDocumentFailure(Runnable test) throws Exception {
    try (MockedConstruction<Workbook> mockedDocumentConstructor = Mockito.mockConstruction(Workbook.class,
        (mock, context) -> {
          throw new RuntimeException("Workbook creation failed");
        })) {
      test.run();
    }
  }
}
