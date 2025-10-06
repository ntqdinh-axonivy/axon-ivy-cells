package com.axonivy.utils.axon.ivy.cells.service;

/**
 * Exception thrown when Spreadsheet conversion operations fail. This is a runtime
 * exception that wraps underlying conversion errors.
 */
public class SpreadsheetConversionException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  /**
   * Constructs a new SpreadsheetConversionException with the specified detail
   * message.
   * 
   * @param message the detail message
   */
  public SpreadsheetConversionException(String message) {
    super(message);
  }

  /**
   * Constructs a new SpreadsheetConversionException with the specified detail
   * message and cause.
   * 
   * @param message the detail message
   * @param cause   the cause
   */
  public SpreadsheetConversionException(String message, Throwable cause) {
    super(message, cause);
  }

  /**
   * Constructs a new SpreadsheetConversionException with the specified cause.
   * 
   * @param cause the cause
   */
  public SpreadsheetConversionException(Throwable cause) {
    super(cause);
  }
}
