package it.docx4j.utils;

import com.topologi.diffx.config.WhiteSpaceProcessing;
import it.docx4j.style.IDocxObjectStyle;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTColumns;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.CTTblPrBase;
import org.docx4j.wml.CTVerticalJc;
import org.docx4j.wml.Color;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcMar;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.U;
import org.docx4j.wml.UnderlineEnumeration;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.List;

public class StyleFactory {

	public static R textFormatting(String textToInsert, IDocxObjectStyle style) {
		return textFormatting(textToInsert, style.getFontColor(), style.getFontSize(), style.getFontFamily(), style.isBold(), style.isUnderline(), style.isItalic());
	}

	public static R textFormatting(String textToInsert, String fontColor, String fontSize, String fontFamily, Boolean bold, Boolean underline, Boolean italic) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		R r = factory.createR();
		RPr rpr = factory.createRPr();

		setFontColor(rpr, fontColor != null && fontColor.length() == 6 ? fontColor : "000000");
		setFontFamily(rpr, fontFamily);
		setFontSize(rpr, fontSize);
		if (bold)
			addBoldStyle(rpr);
		if (underline)
			addUnderlineStyle(rpr);
		if (italic)
			addItalicStyle(rpr);

		r.setRPr(rpr);
		Text text = factory.createText();
		text.setValue(textToInsert);
		text.setSpace(WhiteSpaceProcessing.PRESERVE.name().toLowerCase());
		r.getContent().add(text);
		return r;
	}

	public static void overrideDefaultStyle(MainDocumentPart MainDocumentPart, IDocxObjectStyle styling, String styleId) throws Docx4JException {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Styles styles = MainDocumentPart.getStyleDefinitionsPart().getContents();
		for (Style s : styles.getStyle()) {
			if (s.getStyleId().toLowerCase().equals(styleId)) {
				RPr rpr = s.getRPr();
				if (rpr == null) {
					rpr = factory.createRPr();
					s.setRPr(rpr);
				}
				setFontColor(rpr, styling.getFontColor() != null && styling.getFontColor().length() == 6 ? styling.getFontColor() : "000000");
				setFontFamily(rpr, styling.getFontFamily());
				setFontSize(rpr, styling.getFontSize());
				if (styling.isBold())
					addBoldStyle(rpr);
				if (styling.isUnderline())
					addUnderlineStyle(rpr);
				if (styling.isItalic())
					addItalicStyle(rpr);
			}
		}
	}

	private static void setCellBorders(Tc tableCell, boolean borderTop, boolean borderRight, boolean borderBottom, boolean borderLeft, String borderColor) {
		TcPr tableCellProperties = tableCell.getTcPr();
		if (tableCellProperties == null) {
			tableCellProperties = new TcPr();
			tableCell.setTcPr(tableCellProperties);
		}

		CTBorder border = new CTBorder();
		border.setColor(borderColor);
		border.setSz(new BigInteger("10"));
		border.setVal(STBorder.SINGLE);
		TcPrInner.TcBorders borders = new TcPrInner.TcBorders();
		if (borderBottom) {
			borders.setBottom(border);
		}
		if (borderTop) {
			borders.setTop(border);
		}
		if (borderLeft) {
			borders.setLeft(border);
		}
		if (borderRight) {
			borders.setRight(border);
		}
		tableCellProperties.setTcBorders(borders);
	}

	protected static void setCellWidth(Tc tableCell, int width) {
		if (width > 0) {
			TcPr tableCellProperties = tableCell.getTcPr();
			if (tableCellProperties == null) {
				tableCellProperties = new TcPr();
				tableCell.setTcPr(tableCellProperties);
			}
			TblWidth tableWidth = new TblWidth();
			tableWidth.setType("dxa");
			tableWidth.setW(BigInteger.valueOf(width));
			tableCellProperties.setTcW(tableWidth);
		}
	}

	private static void setCellNoWrap(Tc tableCell) {
		TcPr tableCellProperties = tableCell.getTcPr();
		if (tableCellProperties == null) {
			tableCellProperties = new TcPr();
			tableCell.setTcPr(tableCellProperties);
		}
		BooleanDefaultTrue b = new BooleanDefaultTrue();
		b.setVal(true);
		tableCellProperties.setNoWrap(b);
	}

	protected static void setCellVMerge(Tc tableCell, String mergeVal) {
		if (mergeVal != null) {
			TcPr tableCellProperties = tableCell.getTcPr();
			if (tableCellProperties == null) {
				tableCellProperties = new TcPr();
				tableCell.setTcPr(tableCellProperties);
			}
			TcPrInner.VMerge merge = new TcPrInner.VMerge();
			if (!"close".equals(mergeVal)) {
				merge.setVal(mergeVal);
			}
			tableCellProperties.setVMerge(merge);
		}
	}

	protected static void setCellHMerge(Tc tableCell, int horizontalMergedCells) {
		if (horizontalMergedCells > 1) {
			TcPr tableCellProperties = tableCell.getTcPr();
			if (tableCellProperties == null) {
				tableCellProperties = new TcPr();
				tableCell.setTcPr(tableCellProperties);
			}

			TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
			gridSpan.setVal(new BigInteger(String.valueOf(horizontalMergedCells)));

			tableCellProperties.setGridSpan(gridSpan);
			tableCell.setTcPr(tableCellProperties);
		}
	}

	private static void setCellColor(Tc tableCell, String color) {
		if (color != null) {
			TcPr tableCellProperties = tableCell.getTcPr();
			if (tableCellProperties == null) {
				tableCellProperties = new TcPr();
				tableCell.setTcPr(tableCellProperties);
			}
			CTShd shd = new CTShd();
			shd.setFill(color);
			tableCellProperties.setShd(shd);
		}
	}

	private static void setCellMargins(Tc tableCell, int top, int right, int bottom, int left) {
		TcPr tableCellProperties = tableCell.getTcPr();
		if (tableCellProperties == null) {
			tableCellProperties = new TcPr();
			tableCell.setTcPr(tableCellProperties);
		}
		TcMar margins = new TcMar();

		if (bottom > 0) {
			TblWidth bW = new TblWidth();
			bW.setType("dxa");
			bW.setW(BigInteger.valueOf(bottom));
			margins.setBottom(bW);
		}

		if (top > 0) {
			TblWidth tW = new TblWidth();
			tW.setType("dxa");
			tW.setW(BigInteger.valueOf(top));
			margins.setTop(tW);
		}

		if (left > 0) {
			TblWidth lW = new TblWidth();
			lW.setType("dxa");
			lW.setW(BigInteger.valueOf(left));
			margins.setLeft(lW);
		}

		if (right > 0) {
			TblWidth rW = new TblWidth();
			rW.setType("dxa");
			rW.setW(BigInteger.valueOf(right));
			margins.setRight(rW);
		}

		tableCellProperties.setTcMar(margins);
	}

	private static void setVerticalAlignment(Tc tableCell, STVerticalJc align) {
		if (align != null) {
			TcPr tableCellProperties = tableCell.getTcPr();
			if (tableCellProperties == null) {
				tableCellProperties = new TcPr();
				tableCell.setTcPr(tableCellProperties);
			}

			CTVerticalJc valign = new CTVerticalJc();
			valign.setVal(align);

			tableCellProperties.setVAlign(valign);
		}
	}

	private static void setFontSize(RPr runProperties, String fontSize) {
		if (fontSize != null && !fontSize.isEmpty()) {
			HpsMeasure size = new HpsMeasure();
			size.setVal(BigInteger.valueOf(Integer.parseInt(fontSize) * 2));
			runProperties.setSz(size);
			runProperties.setSzCs(size);
		}
	}

	private static void setFontFamily(RPr runProperties, String fontFamily) {
		if (fontFamily != null) {
			RFonts rf = runProperties.getRFonts();
			if (rf == null) {
				rf = new RFonts();
				runProperties.setRFonts(rf);
			}
			rf.setAscii(fontFamily);
			rf.setHAnsi(fontFamily);
			rf.setAsciiTheme(null);
			rf.setHAnsiTheme(null);
		}
	}

	private static void setFontColor(RPr runProperties, String color) {
		if (color != null) {
			Color c = new Color();
			c.setVal(color);
			runProperties.setColor(c);
		}
	}

	private static void setHorizontalAlignment(P paragraph, JcEnumeration hAlign) {
		if (hAlign != null) {
			PPr pprop = new PPr();
			Jc align = new Jc();
			align.setVal(hAlign);
			pprop.setJc(align);
			paragraph.setPPr(pprop);
		}
	}

	private static void addBoldStyle(RPr runProperties) {
		BooleanDefaultTrue b = new BooleanDefaultTrue();
		b.setVal(true);
		runProperties.setB(b);
	}

	private static void addItalicStyle(RPr runProperties) {
		BooleanDefaultTrue b = new BooleanDefaultTrue();
		b.setVal(true);
		runProperties.setI(b);
	}

	private static void addUnderlineStyle(RPr runProperties) {
		U val = new U();
		val.setVal(UnderlineEnumeration.SINGLE);
		runProperties.setU(val);
	}

	public static void addStyle(StyleDefinitionsPart styleDefinitionsPart, String styleName) throws Docx4JException {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Style myNewStyle = Context.getWmlObjectFactory().createStyle();
		myNewStyle.setType("paragraph");
		myNewStyle.setStyleId(styleName);
		Style.Name n = Context.getWmlObjectFactory().createStyleName();
		n.setVal(styleName);
		myNewStyle.setName(n);
		styleDefinitionsPart.getContents().getStyle().add(myNewStyle);
		Style.BasedOn based = Context.getWmlObjectFactory().createStyleBasedOn();
		based.setVal("Normal");
		myNewStyle.setBasedOn(based);
		RPr rpr = myNewStyle.getRPr();
		if (rpr == null) {
			rpr = factory.createRPr();
			myNewStyle.setRPr(rpr);
		}
		setFontColor(rpr, "f000f0");
		addBoldStyle(rpr);
		setFontFamily(rpr, "Algerian");
	}


	protected static void addTableCell(Tr tableRow, P image, int width, int horizontalMergedCells, String verticalMergedVal, IDocxObjectStyle style) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Tc tableCell = factory.createTc();
		addImageCellStyle(tableCell, image, style);
		setCellWidth(tableCell, width);
		setCellVMerge(tableCell, verticalMergedVal);
		setCellHMerge(tableCell, horizontalMergedCells);
		tableRow.getContent().add(tableCell);
	}

	protected static void addTableCell(Tr tableRow, String content, int width, int horizontalMergedCells, String verticalMergedVal, IDocxObjectStyle style) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Tc tableCell = factory.createTc();
		addCellStyle(tableCell, content, style);
		setCellWidth(tableCell, width);
		setCellVMerge(tableCell, verticalMergedVal);
		setCellHMerge(tableCell, horizontalMergedCells);
		if (style.isNoWrap()) {
			setCellNoWrap(tableCell);
		}
		tableRow.getContent().add(tableCell);
	}

	protected static void addCellStyle(Tc tableCell, String content, IDocxObjectStyle style) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();
		Text text = factory.createText();
		text.setValue(content);
		R run = factory.createR();
		run.getContent().add(text);
		paragraph.getContent().add(run);
		if (style != null) {
			setHorizontalAlignment(paragraph, style.getHorizAlignment());
			RPr runProperties = factory.createRPr();
			if (style.isBold()) {
				addBoldStyle(runProperties);
			}
			if (style.isItalic()) {
				addItalicStyle(runProperties);
			}
			if (style.isUnderline()) {
				addUnderlineStyle(runProperties);
			}
			setFontSize(runProperties, style.getFontSize());
			setFontColor(runProperties, style.getFontColor());
			setFontFamily(runProperties, style.getFontFamily());

			setCellMargins(tableCell, style.getTop(), style.getRight(), style.getBottom(), style.getLeft());
			setCellColor(tableCell, style.getBackground());
			setVerticalAlignment(tableCell, style.getVerticalAlignment());
			setCellBorders(tableCell, style.isBorderTop(), style.isBorderRight(), style.isBorderBottom(), style.isBorderLeft(), style.getBorderColor());
			run.setRPr(runProperties);
		}
		tableCell.getContent().add(paragraph);
	}

	protected static void addImageCellStyle(Tc tableCell, P image, IDocxObjectStyle style) {
		if (style != null) {
			setCellMargins(tableCell, style.getTop(), style.getRight(), style.getBottom(), style.getLeft());
			setCellColor(tableCell, style.getBackground());
			setVerticalAlignment(tableCell, style.getVerticalAlignment());
			setHorizontalAlignment(image, style.getHorizAlignment());
			setCellBorders(tableCell, style.isBorderTop(), style.isBorderRight(), style.isBorderBottom(), style.isBorderLeft(), style.getBorderColor());
		}
		tableCell.getContent().add(image);
	}

	public static P newImage(WordprocessingMLPackage wordMLPackage, byte[] bytes, String filenameHint, String altText, int id1, int id2, long cx) throws Exception {
		BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
		Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, cx, false);

		ObjectFactory factory = Context.getWmlObjectFactory();

		Drawing drawing = factory.createDrawing();
		R run = factory.createR();
		run.getContent().add(drawing);
		P p = factory.createP();
		p.getContent().add(run);

		drawing.getAnchorOrInline().add(inline);
		return p;
	}

	public static byte[] getImageBytes(String ImageName) throws IOException {
		File file = new File(ImageName);
		InputStream inputStream = new java.io.FileInputStream(file);
		long fileLength = file.length();
		byte[] bytes = new byte[(int) fileLength];

		int offset = 0;
		int numRead;
		while (offset < bytes.length
				&& (numRead = inputStream.read(bytes, offset, bytes.length - offset)) >= 0) {
			offset += numRead;
		}
		inputStream.close();

		return bytes;
	}

	public static Tbl createTableWithContent(IDocxObjectStyle headerStyle, IDocxObjectStyle evenStyle, IDocxObjectStyle oddStyle, List<String> headerValue, List<Object> cellValue) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Tbl table = factory.createTbl();

		Tr tableRow = factory.createTr();
		int column = headerValue.size();
		for (String value : headerValue) {
			addTableCell(tableRow, value.toUpperCase(), 0, 0, "restart", headerStyle);
		}
		table.getContent().add(tableRow);
		int row = 0;
		for (int i = 0; i < cellValue.size(); i++) {
			if (i == 0 || (i % column == 0 && i != 1)) {
				tableRow = factory.createTr();
				table.getContent().add(tableRow);
				row++;
			}
			Object obj = cellValue.get(i);
			if (obj instanceof String) {
				addTableCell(tableRow, obj.toString(), 0, 0, "restart", (row % 2 == 0) ? evenStyle : oddStyle);
			} else if (obj instanceof P) {
				addTableCell(tableRow, (P) obj, 0, 0, "restart", (row % 2 == 0) ? evenStyle : oddStyle);
			} else if (obj instanceof P.Hyperlink) {
				P p4 = factory.createP();
				p4.getContent().add(obj);
				addTableCell(tableRow, p4, 0, 0, "restart", (row % 2 == 0) ? evenStyle : oddStyle);
			}
		}
		//Restart e close per gestire il merge
		TblPr tblPr = new TblPr();
		CTTblPrBase.TblStyle tblStyle = new CTTblPrBase.TblStyle();
		tblStyle.setVal("TableGrid");
		tblPr.setTblStyle(tblStyle);
		table.setTblPr(tblPr);

		return table;
	}

	public static P.Hyperlink createHyperlink(MainDocumentPart mdp, String url, String text) {
		try {
			org.docx4j.relationships.ObjectFactory factory = new org.docx4j.relationships.ObjectFactory();

			org.docx4j.relationships.Relationship rel = factory.createRelationship();
			rel.setType(Namespaces.HYPERLINK);
			rel.setTarget(url);
			rel.setTargetMode("External");
			mdp.getRelationshipsPart().addRelationship(rel);
			String hpl = "<w:hyperlink r:id=\"" + rel.getId() + "\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
					"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" >" +
					"<w:r>" +
					"<w:rPr>" +
					"<w:rStyle w:val=\"Hyperlink\" />" +
					"</w:rPr>" +
					"<w:t>" + text + "</w:t>" +
					"</w:r>" +
					"</w:hyperlink>";
			return (P.Hyperlink) XmlUtils.unmarshalString(hpl);

		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	public static PPr customizeLayoutSection(int column) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		PPr ppr = factory.createPPr();
		SectPr sectPr = factory.createSectPr();
		CTColumns ctColumns = factory.createCTColumns();
		ctColumns.setNum(BigInteger.valueOf(column));
		sectPr.setCols(ctColumns);
		ppr.setSectPr(sectPr);
		return ppr;
	}

	public static PPr createSectionBreak() {
		ObjectFactory factory = Context.getWmlObjectFactory();
		PPr ppr = factory.createPPr();
		SectPr sectPr = factory.createSectPr();
		SectPr.Type sectPrType = new SectPr.Type();
		sectPrType.setVal("continuous");
		sectPr.setType(sectPrType);
		BigInteger bi = new BigInteger("9");
		SectPr.PgSz xx = factory.createSectPrPgSz();
		xx.setCode(bi);
		sectPr.setPgSz(xx);
		ppr.setSectPr(sectPr);
		return ppr;
	}

}
