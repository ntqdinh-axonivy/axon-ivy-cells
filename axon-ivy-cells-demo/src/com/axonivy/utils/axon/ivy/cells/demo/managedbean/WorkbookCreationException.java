package com.axonivy.utils.axon.ivy.cells.demo.managedbean;

/**
 * Exception thrown when workbook creation or manipulation operations fail.
 * This exception provides specific error handling for workbook-related operations.
 */
public class WorkbookCreationException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  /**
   * Constructs a new WorkbookCreationException with the specified detail message.
   * 
   * @param message the detail message
   */
  public WorkbookCreationException(String message) {
    super(message);
  }

  /**
   * Constructs a new WorkbookCreationException with the specified detail message and cause.
   * 
   * @param message the detail message
   * @param cause the cause of the exception
   */
  public WorkbookCreationException(String message, Throwable cause) {
    super(message, cause);
  }

  /**
   * Constructs a new WorkbookCreationException with the specified cause.
   * 
   * @param cause the cause of the exception
   */
  public WorkbookCreationException(Throwable cause) {
    super(cause);
  }
}