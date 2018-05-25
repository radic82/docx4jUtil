package it.docx4j.item;

import it.docx4j.style.IDocxObjectStyle;

/**
 * Class to manage image
 */
public class DocImg extends DocItem {

	private String path;
	private String fileNameHint;
	private String altText;

	public DocImg(String path) {
		this(path, null);
	}

	public DocImg(String path, IDocxObjectStyle style) {
		this.path = path;
		this.style = style;
		this.fileNameHint = null;
		this.altText = null;
	}

	public String getPath() {
		return path;
	}

	public String getFileNameHint() {
		return fileNameHint;
	}

	public void setFileNameHint(String fileNameHint) {
		this.fileNameHint = fileNameHint;
	}

	public String getAltText() {
		return altText;
	}

	public void setAltText(String altText) {
		this.altText = altText;
	}
}