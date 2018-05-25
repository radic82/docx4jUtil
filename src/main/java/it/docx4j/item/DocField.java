package it.docx4j.item;

import it.docx4j.style.custom.DocxStyleEnum;

/**
 * Class to manage simple Item with title and value
 */
public class DocField extends DocItem {
	public static final java.lang.CharSequence SEPARATOR = " ";
	private DocText title;
	private Object value;

	/**
	 * Default Implementation
	 *
	 * @param title title
	 * @param value value
	 */
	public DocField(String title, Object value) {
		this.title = new DocText(title, DocxStyleEnum.TEXT_BOLD);
		this.value = value;
		this.setStyle(DocxStyleEnum.TEXT_NORMAL);
	}

	public DocText getTitle() {
		return title;
	}

	public Object getValue() {
		return value;
	}
}
