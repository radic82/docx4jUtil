package it.docx4j.item;

import it.docx4j.style.custom.DocxStyleEnum;

/**
 * Class to manage link
 */
public class DocLink extends DocItem {

	private String url;
	private String text;

	public DocLink(String url, String text) {
		this(url, text, DocxStyleEnum.TEXT_NORMAL);
	}

	public DocLink(String url, String text, DocxStyleEnum style) {
		this.url = url;
		this.text = text;
		this.style = style;
	}

	public String getUrl() {
		return url;
	}

	public String getText() {
		return text;
	}
}