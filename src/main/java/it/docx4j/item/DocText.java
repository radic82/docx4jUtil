package it.docx4j.item;


import it.docx4j.style.IDocxObjectStyle;

/**
 * Class to manage simple Item with only text (example: paragraph)
 */
public class DocText extends DocItem {

	private String text;

	public DocText(String text) {
		this.text = text;
	}

	public DocText(String text, IDocxObjectStyle style) {
		this(text);
		this.style = style;
	}

	public DocText(String text, String styleId) {
		this(text);
		this.styleId = styleId;
	}

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}
}
