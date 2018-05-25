package it.docx4j.section;

import it.docx4j.item.DocText;
import it.docx4j.style.IDocxObjectStyle;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public abstract class DocSimpleSection<T> {
	protected DocText title;
	protected T content;

	public DocSimpleSection(DocText title, T content) {
		this.title = title;
		this.content = content;
	}

	public DocSimpleSection(T content) {
		this(null, content);
	}

	protected T getContent() {
		return content;
	}

	protected void setContent(T content) {
		this.content = content;
	}

	protected boolean hasContent() {
		return this.getContent() != null;
	}

	protected boolean hasValidContent() {
		return false;
	}

	public DocText getTitle() {
		return title;
	}

	public String getTitleText() {
		if (!this.hasContent() || this.hasValidContent())
			if (this.title != null)
				return this.title.getText();
		return null;
	}

	public IDocxObjectStyle getTitleStyle() {
		if (this.title != null)
			return this.title.getStyle();
		return null;
	}

	public String getTitleStyleId() {
		if (this.title != null)
			return this.title.getStyleId();
		return null;
	}

	public <K> K retrieveObject(WordprocessingMLPackage wordMLPackage) {
		return null;
	}
}