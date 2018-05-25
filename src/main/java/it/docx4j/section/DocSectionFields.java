package it.docx4j.section;

import it.docx4j.item.DocField;
import it.docx4j.item.DocText;
import it.docx4j.utils.SectionFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;

import java.util.List;

/**
 * Class to manages section with title and list of key - value
 */
public class DocSectionFields extends DocSimpleSection<List<DocField>> {

	public DocSectionFields(DocText title, List<DocField> content) {
		super(title, content);
	}

	public DocSectionFields(String title, List<DocField> content) {
		super(new DocText(title, "Heading1"), content);
	}

	@Override
	protected boolean hasValidContent() {
		return this.hasContent() && (this.getContent() != null && !this.getContent().isEmpty());
	}

	@Override
	public List<DocField> getContent() {
		return this.content;
	}

	@Override
	public void setContent(List<DocField> content) {
		this.content = content;
	}

	@Override
	public List<P> retrieveObject(WordprocessingMLPackage wordMLPackage) {
		if (this.hasValidContent())
			return SectionFactory.fromFieldListToParagraphList(this.getContent(), wordMLPackage);
		return null;
	}
}
