package com.axonivy.utils.axon.ivy.cells.demo.managedbean;

import java.io.ByteArrayInputStream;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.ViewScoped;

import org.primefaces.model.DefaultStreamedContent;
import org.primefaces.model.file.UploadedFile;

import com.axonivy.utils.axon.ivy.cells.service.CellFactory;

import ch.ivyteam.ivy.environment.Ivy;

@ManagedBean
@ViewScoped
public class CellFactoryBean {
	private static final long serialVersionUID = 1L;
	private static final String PDF_EXTENSION = ".pdf";
	private static final String DOT = ".";
	private UploadedFile uploadedFile;
	private DefaultStreamedContent convertedFile;

	public void convert() {
		if (uploadedFile != null) {
			Ivy.log().warn(CellFactory.convert().from(uploadedFile.getContent()).toPdf().asBytes());
			String pdfFileName = updateFileExtension();
			setConvertedFile(DefaultStreamedContent.builder().name(pdfFileName).contentType("application/pdf")
					.stream(() -> new ByteArrayInputStream(
							CellFactory.convert().from(uploadedFile.getContent()).toPdf().asBytes()))
					.build());
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
}
