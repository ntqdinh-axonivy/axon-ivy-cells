package com.axonivy.utils.axon.ivy.cells.demo.managedbean;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.ViewScoped;

import org.apache.commons.lang.StringUtils;
import org.primefaces.model.DefaultStreamedContent;
import org.primefaces.model.file.UploadedFile;

import com.aspose.cells.Chart;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.axonivy.utils.axon.ivy.cells.service.CellFactory;

import ch.ivyteam.ivy.environment.Ivy;

@ManagedBean
@ViewScoped
public class CellFactoryBean {
  private static final String PDF_EXTENSION = ".pdf";
  private static final String DOT = ".";
  private UploadedFile uploadedFile;
  private DefaultStreamedContent convertedFile;
  private String cellToUpdate;
  private double newCellValue;
  private int workingSheetIndex;

  public void convert() {
    if (uploadedFile != null) {
      String pdfFileName = updateFileExtension();
      setConvertedFile(DefaultStreamedContent.builder().name(pdfFileName).contentType("application/pdf")
          .stream(
              () -> new ByteArrayInputStream(CellFactory.convert().from(uploadedFile.getContent()).toPdf().asBytes()))
          .build());
    }
  }

  /**
   * Updates the existing uploaded workbook, modifies the specified cell with a
   * new value, refreshes all charts and diagrams in the workbook, and returns it
   * as a PDF stream. This method works with the uploaded XLSX file.
   */
  public void updateWorkbook() {
    if (uploadedFile == null) {
      throw new WorkbookCreationException("No file uploaded. Please upload a workbook file first.");
    }

    try {
      Workbook workbook = CellFactory.get(() -> {
        try {
          return new Workbook(new ByteArrayInputStream(uploadedFile.getContent()));
        } catch (Exception e) {
          throw new WorkbookCreationException("Failed to load uploaded workbook", e);
        }
      });

      WorksheetCollection worksheets = workbook.getWorksheets();
      Worksheet worksheet = worksheets.get(workingSheetIndex);

      if (StringUtils.isNotBlank(cellToUpdate)) {
        worksheet.getCells().get(cellToUpdate).putValue(newCellValue);
      }

      refreshChartsAndDiagrams(workbook);

      // Convert to PDF instead of XLSX
      ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
      workbook.save(outputStream, SaveFormat.PDF);
      String fileName = getUpdatedPdfFileName();
      setConvertedFile(DefaultStreamedContent.builder().name(fileName).contentType("application/pdf")
          .stream(() -> new ByteArrayInputStream(outputStream.toByteArray())).build());
    } catch (Exception e) {
      Ivy.log().error("Failed to update workbook", e);
      throw new WorkbookCreationException("Failed to update workbook", e);
    }
  }

  /**
   * Generates an updated PDF filename based on the original uploaded file name.
   */
  private String getUpdatedPdfFileName() {
    String originalName = uploadedFile.getFileName();
    if (originalName != null && originalName.contains(DOT)) {
      String baseName = originalName.substring(0, originalName.lastIndexOf(DOT));
      return baseName + "_updated" + PDF_EXTENSION;
    }
    return "updated_workbook.pdf";
  }

  /**
   * Refreshes all charts and diagrams in the workbook to reflect updated data.
   * This ensures that any visual elements are synchronized with the current data.
   */
  private void refreshChartsAndDiagrams(Workbook workbook) {
    try {
      // Iterate through all worksheets
      for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
        Worksheet worksheet = workbook.getWorksheets().get(i);

        // Refresh all charts in the worksheet
        for (int j = 0; j < worksheet.getCharts().getCount(); j++) {
          Chart chart = worksheet.getCharts().get(j);
          chart.calculate();
        }

        // Refresh all pivot tables in the worksheet
        for (int k = 0; k < worksheet.getPivotTables().getCount(); k++) {
          worksheet.getPivotTables().get(k).calculateData();
        }
      }
      workbook.calculateFormula();
    } catch (Exception e) {
      Ivy.log().warn("Failed to refresh charts and diagrams", e);
    }
  }

  private String updateFileExtension() {
    String originalName = uploadedFile.getFileName();
    String baseName = originalName != null && originalName.contains(DOT)
        ? originalName.substring(0, originalName.lastIndexOf(DOT))
        : originalName;
    return baseName + PDF_EXTENSION;
  }

  public DefaultStreamedContent getConvertedFile() {
    return convertedFile;
  }

  public void setConvertedFile(DefaultStreamedContent convertedFile) {
    this.convertedFile = convertedFile;
  }

  public UploadedFile getUploadedFile() {
    return uploadedFile;
  }

  public void setUploadedFile(UploadedFile uploadedFile) {
    this.uploadedFile = uploadedFile;
  }

  public String getCellToUpdate() {
    return cellToUpdate;
  }

  public void setCellToUpdate(String cellToUpdate) {
    this.cellToUpdate = cellToUpdate;
  }

  public double getNewCellValue() {
    return newCellValue;
  }

  public void setNewCellValue(double newCellValue) {
    this.newCellValue = newCellValue;
  }

  public int getWorkingSheetIndex() {
    return workingSheetIndex;
  }

  public void setWorkingSheetIndex(int workingSheetIndex) {
    this.workingSheetIndex = workingSheetIndex;
  }
}
