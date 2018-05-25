package it.docx4j.section;

import it.docx4j.item.DocTable;
import it.docx4j.item.DocText;
import it.docx4j.utils.SectionFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Tbl;

/**
 * Class to manages section with title and table
 */
public class DocSectionTable extends DocSimpleSection<DocTable> {
	/**
	 * Default
	 *
	 * @param title   section title
	 * @param content
	 */
	public DocSectionTable(String title, DocTable content) {
		super(new DocText(title, "Heading1"), content);

	}

	public DocSectionTable(DocTable content) {
		super(content);
	}

	@Override
	protected boolean hasValidContent() {
		if (this.hasContent())
			if (this.getContent().getRowSize() > 1)
				if (this.getContent().getColumnSize() > 0)
					return true;
		return false;
	}

	@Override
	public DocTable getContent() {
		return this.content;
	}

	@Override
	public void setContent(DocTable content) {
		this.content = content;
	}

	@Override
	public Tbl retrieveObject(WordprocessingMLPackage wordMLPackage) {
		if (this.hasValidContent())
			return SectionFactory.fromDocTableToTable(wordMLPackage, this.getContent());
		return null;
	}
}
