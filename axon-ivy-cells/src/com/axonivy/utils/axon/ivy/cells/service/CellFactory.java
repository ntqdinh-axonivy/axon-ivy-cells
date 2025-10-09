package com.axonivy.utils.axon.ivy.cells.service;

import java.io.InputStream;
import java.util.function.Supplier;

import com.aspose.cells.License;

import ch.ivyteam.ivy.ThirdPartyLicenses;
import ch.ivyteam.ivy.environment.Ivy;

public class CellFactory {
  private static License license;

  private CellFactory() {
  }

  static {
    loadLicense();
  }

  /**
   * Initializes the Aspose CellFactory license.
   * <p>
   * Ensures the license is loaded once per request. If not already set, this
   * method retrieves the license from {@link ThirdPartyLicenses} and applies it
   * to the Aspose {@link License} instance.
   * </p>
   *
   * <p>
   * In case of failure, the exception is logged and the license reference is
   * reset to {@code null}, leaving the application in evaluation mode.
   * </p>
   */
  public static void loadLicense() {
    if (license != null) {
      return;
    }
    try {
      InputStream in = ThirdPartyLicenses.getDocumentFactoryLicense();
      if (in != null) {
        license = new License();
        license.setLicense(in);
      }
    } catch (Exception e) {
      Ivy.log().error(e);
      license = null;
    }
  }

  /**
   * Creates a new document converter for fluent API usage.
   * <p>
   * Usage examples:
   * 
   * <pre>
   * // Convert to PDF as bytes
   * byte[] pdfBytes = WordFactory.convert().from(file).toPdf().asBytes();
   * 
   * // Convert to any format as file
   * File outputFile = WordFactory.convert().from(file).to(Format.PDF).asFile("/path/to/output.pdf");
   * </pre>
   * </p>
   * 
   * @return a new DocumentConverter instance
   */
  public static SpreadsheetConverter convert() {
    return new SpreadsheetConverter();
  }

  /**
   * Executes a supplier function after ensuring the Aspose CellFactory
   * license is loaded.
   * <p>
   * This method guarantees that the license is initialized before invoking the
   * provided {@link Supplier}. It allows callers to transparently execute logic
   * that depends on a valid license, without duplicating license initialization
   * checks.
   * </p>
   *
   * @param supplier the function to execute
   * @param <T>      the return type of the supplier
   * @return the result produced by the supplier
   */
  public static <T> T get(Supplier<T> supplier) {
    return supplier.get();
  }

  /**
   * Executes a runnable task after ensuring the Aspose CellFactory license is
   * loaded.
   * <p>
   * This method guarantees that the license is initialized before invoking the
   * provided {@link Runnable}. It allows callers to run license-dependent
   * operations in a safe and consistent manner.
   * </p>
   *
   * @param run the task to execute
   */
  public static void run(Runnable run) {
    run.run();
  }
}
