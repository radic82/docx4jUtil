package it.docx4j.section;

import it.docx4j.item.DocText;

/**
 * Class to manages section with only title
 */
public class DocSectionTitle extends DocSimpleSection<DocText> {

	/**
	 * Default Implementation
	 *
	 * @param title
	 */
	public DocSectionTitle(String title) {
		super(new DocText(title, "Title"), null);
	}

	public DocSectionTitle(DocText title) {
		super(title, null);
	}
}
