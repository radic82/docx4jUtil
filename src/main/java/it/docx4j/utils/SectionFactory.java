package it.docx4j.utils;

import it.docx4j.item.DocField;
import it.docx4j.item.DocImg;
import it.docx4j.item.DocLink;
import it.docx4j.item.DocRow;
import it.docx4j.item.DocTable;
import it.docx4j.item.DocText;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Class to manage several type of sections
 */
public class SectionFactory extends StyleFactory {

	/**
	 * @param wordMLPackage
	 * @param docTable
	 * @return
	 */
	public static Tbl fromDocTableToTable(WordprocessingMLPackage wordMLPackage, DocTable docTable) {
		if (docTable.getColumnSize() <= 0)
			return null;

		int writableWidthTwips = wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
		Tbl table = TblFactory.createTable(docTable.getRowSize(), docTable.getColumnSize(), writableWidthTwips / docTable.getColumnSize());
		//table.getTblPr().getTblStyle().setVal("ListTable1Light-Accent5");

		for (int i = 0; i < docTable.getRowSize(); i++) {
			DocRow docRow = docTable.getRows().get(i);
			//docRow.setStyle(null);
			Tr tableRow = (Tr) table.getContent().get(i);
			for (int j = 0; j < docTable.getColumnSize(); j++) {
				Object value = docRow.getRow().get(j);
				Tc tc = (Tc) tableRow.getContent().get(j);
				tc.getContent().clear();
				if (value instanceof String) {
					//populateP((P) tc.getContent().get(0), (docRow.isHeader() ? String.valueOf(value).toUpperCase() : String.valueOf(value)));
					addCellStyle(tc, (docRow.isHeader() ? String.valueOf(value).toUpperCase() : String.valueOf(value)), docRow.getStyle());
				} else if (value instanceof DocText) {
					addImageCellStyle(tc, fromDocTextToParagraph((DocText) value), docRow.getStyle());
				} else if (value instanceof DocImg) {
					addImageCellStyle(tc, fromDocImgToParagraph((DocImg) value, wordMLPackage), docRow.getStyle());
				} else if (value instanceof DocLink) {
					addImageCellStyle(tc, fromDocLinkToParagraph((DocLink) value, wordMLPackage), docRow.getStyle());
				}
				setCellWidth(tc, docRow.isHeader() ? String.valueOf(value).length() : 0);
				setCellHMerge(tc, 0);
				setCellVMerge(tc, "restart");
			}
		}
		return table;
	}

	/**
	 * @param fields
	 * @return
	 */
	public static List<P> fromFieldListToParagraphList(List<DocField> fields, WordprocessingMLPackage wordMLPackage) {
		return fields.stream().map(field -> fromDocFieldToParagrah(field, wordMLPackage)).collect(Collectors.toList());
	}

	/**
	 * @param simpleText
	 * @return
	 */
	public static P fromDocTextToParagraph(DocText simpleText) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();
		fromDocTextToParagraph(simpleText, paragraph);
		return paragraph;
	}


	private static P fromDocFieldToParagrah(DocField field, WordprocessingMLPackage wordMLPackage) {
		if (field == null && field.getTitle() == null)
			return null;
		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();
		fromDocTextToParagraph(field.getTitle(), paragraph);
		Object value = field.getValue();
		if (value != null) {
			//paragraph.getContent().add(textFormatting(" ", DocxStyleEnum.TEXT_NORMAL));
			if (value instanceof List) {
				paragraph.getContent().add(textFormatting(Arrays.stream(((List) value).toArray())
						.map(item -> (CharSequence) item)
						.collect(Collectors.joining(DocField.SEPARATOR)), field.getStyle()));
			} else if (value instanceof DocText) {
				fromDocTextToParagraph((DocText) value, paragraph);
			} else if (value instanceof DocLink) {
				paragraph.getContent().add(createHyperlink(wordMLPackage.getMainDocumentPart()
						, ((DocLink) value).getUrl(), ((DocLink) value).getText()));
			} else if (value instanceof DocImg) {
				fromDocImgToParagraph((DocImg) value, paragraph, wordMLPackage);
			} else {
				paragraph.getContent().add(textFormatting(String.valueOf(value), field.getStyle()));
			}
		}
		return paragraph;
	}

	private static void fromDocTextToParagraph(DocText docText, P paragraph) {
		paragraph.getContent().add(textFormatting(docText.getText() + " ", docText.getStyle()));
	}

	private static P fromDocLinkToParagraph(DocLink docLink, WordprocessingMLPackage wordMLPackage) {
		createHyperlink(wordMLPackage.getMainDocumentPart()
				, docLink.getUrl(), docLink.getText());

		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();
		return paragraph;
	}

	private static P fromDocImgToParagraph(DocImg docImg, WordprocessingMLPackage wordMLPackage) {
		return fromDocImgToParagraph(docImg, null, wordMLPackage);
	}

	private static P fromDocImgToParagraph(DocImg docImg, P paragraph, WordprocessingMLPackage wordMLPackage) {
		try {
			byte[] bytes = StyleFactory.getImageBytes(docImg.getPath());
			if (bytes != null) {
				BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
				Inline inline = imagePart.createImageInline(docImg.getFileNameHint(), docImg.getAltText(), 1, 1, 500, false);

				ObjectFactory factory = Context.getWmlObjectFactory();
				Drawing drawing = factory.createDrawing();
				drawing.getAnchorOrInline().add(inline);

				R run = factory.createR();
				run.getContent().add(drawing);
				if (paragraph == null)
					paragraph = factory.createP();
				paragraph.getContent().add(run);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return paragraph;
	}

	private static void populateP(P paragraph, String content) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		if (paragraph == null)
			paragraph = factory.createP();
		if (paragraph.getContent() == null || paragraph.getContent().isEmpty()) {
			Text text = factory.createText();
			text.setValue(content);
			R run = factory.createR();
			run.getContent().add(text);
			paragraph.getContent().add(run);
		} else {
			R run = (R) paragraph.getContent().get(0);
			if (run.getContent() == null || run.getContent().isEmpty()) {
				Text text = factory.createText();
				text.setValue(content);
				run.getContent().add(text);
			} else {
				Text text = (Text) run.getContent().get(0);
				text.setValue(content);
			}
		}
	}
}
