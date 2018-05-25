package report;

import it.docx4j.style.custom.CustomHeadingStyle;
import it.docx4j.style.custom.CustomNormalStyle;
import it.docx4j.style.custom.CustomTitleStyle;
import it.docx4j.style.custom.DocxStyleEnum;
import it.docx4j.utils.StyleFactory;
import org.apache.commons.lang3.RandomUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.toc.TocGenerator;
import org.docx4j.wml.Br;
import org.docx4j.wml.P;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.Tbl;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class CreateReport {
	private static org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
	private final static String SRC_RESOURCES = "src/main/resources/";
	private final static String SRC_RESOURCES_IMAGE = SRC_RESOURCES + "img/";

	@Test
	public void createReportTextFormat() throws Exception {
		String newFileName = "01.docx";

		File f = new File(SRC_RESOURCES + newFileName);
		if (f.exists())
			f.delete();

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

		P p1 = factory.createP();
		Random rand = new Random();
		for (int i = 0; i < 100; i++) {
			p1.getContent().add(StyleFactory.textFormatting("text ", randomColor(rand), "12", randomFont(), RandomUtils.nextInt(0, 2) == 1, RandomUtils.nextInt(0, 2) == 1, RandomUtils.nextInt(0, 2) == 1));
		}
		log("Customize text in paragraph", true);
		wordMLPackage.getMainDocumentPart().getContent().add(p1);

		for (int i = 0; i < 20; i++) {
			P p2 = factory.createP();
			p2.getContent().add(StyleFactory.textFormatting(i + "text", randomColor(rand), (i * 2 + 12) + "", randomFont(), RandomUtils.nextInt(0, 2) == 1, RandomUtils.nextInt(0, 2) == 1, RandomUtils.nextInt(0, 2) == 1));
			wordMLPackage.getMainDocumentPart().getContent().add(p2);
		}
		log("Customize text paragraph", true);
		P p3 = factory.createP();
		p3.getContent().add(StyleFactory.textFormatting("TITLE RED", DocxStyleEnum.TEXT_TITLE_RED));
		wordMLPackage.getMainDocumentPart().getContent().add(p3);
		log("Customize text paragraph with enum", true);
		P p4 = factory.createP();
		p4.getContent().add(StyleFactory.createHyperlink(wordMLPackage.getMainDocumentPart(), "http://radic.altervista.org", "aa"));
		wordMLPackage.getMainDocumentPart().getContent().add(p4);
		log("Add Link", true);

		byte[] bytes = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "ok.jpg");
		P p5 = StyleFactory.newImage(wordMLPackage, bytes, "caption", "caption", 1, 1, 5000);
		p5.getContent().add(StyleFactory.textFormatting("Test Red Title", DocxStyleEnum.TEXT_NORMAL));
		wordMLPackage.getMainDocumentPart().getContent().add(p5);
		log("Add Image", true);

		wordMLPackage.save(f);
		log(f.getAbsolutePath(), true);
	}

	@Test
	public void createReportStyle() throws Exception {
		String newFileName = "02.docx";

		File f = new File(SRC_RESOURCES + newFileName);
		if (f.exists())
			f.delete();

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "Title with style");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "Heading1 with style");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "Heading2 with style");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "Heading2.1 with style");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "Heading2.2 with style");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Caption", "Caption with style Caption");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Subtitle", "Subtitle");
		wordMLPackage.getMainDocumentPart().addParagraphOfText("Text normal");
		log("create paragraphs with style", true);

		TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
		tocGenerator.generateToc(0, " TOC \\o \"1-3\" \\h \\z \\u ", false);
		log("Generate TOC", true);

		StyleFactory.addStyle(wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart(), "JAStyle");
		log("Add Custom [JAStyle] style", true);

		StyleFactory.overrideDefaultStyle(wordMLPackage.getMainDocumentPart(), new CustomNormalStyle(), "normal");
		log("Override [normal] style", true);

		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("JAStyle", "My text with JAStyle");
		log("Use Custom style", false);

		wordMLPackage.save(f);
		log(f.getAbsolutePath(), true);
	}

	@Test
	public void createTable() throws Exception {
		String newFileName = "03.docx";

		File f = new File(SRC_RESOURCES + newFileName);
		if (f.exists())
			f.delete();

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		byte[] bytes = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "ok.jpg");
		byte[] bytes1 = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "ko.jpg");
		P imageOk = StyleFactory.newImage(wordMLPackage, bytes, "caption", "caption", 1, 1, 500);
		P imageKO = StyleFactory.newImage(wordMLPackage, bytes1, "caption", "caption", 1, 1, 500);
		log("PDF create image", true);
		List<String> header = new ArrayList<>();
		header.add("Who?");
		header.add("Role");
		header.add("OK / KO");
		header.add("Max Speed");
		header.add("Description");

		List<Object> cellValue = new ArrayList<>();
		cellValue.add("Andrea Radice");
		cellValue.add("SW");
		cellValue.add(imageOk);
		cellValue.add("10.000");
		cellValue.add("this is simple text");
		cellValue.add("Marc Morris");
		cellValue.add("ARCH");
		cellValue.add(imageKO);
		cellValue.add("9.000");
		cellValue.add("Bye Bye text test");

		Tbl table = StyleFactory.createTableWithContent(DocxStyleEnum.TABLE_HEADER_CELL, DocxStyleEnum.TABLE_EVEN_CELL, DocxStyleEnum.TABLE_ODD_CELL, header, cellValue);
		wordMLPackage.getMainDocumentPart().addObject(table);
		wordMLPackage.save(f);
		log(f.getAbsolutePath(), true);
	}

	@Test
	public void createReportDoubleColumn() throws Exception {
		String newFileName = "04.docx";

		File f = new File(SRC_RESOURCES + newFileName);
		if (f.exists())
			f.delete();

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.customizeLayoutSection(1));
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "My Firts Custom Report");
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());

		P par = factory.createP();
		par.getContent().add(StyleFactory.customizeLayoutSection(2));

		StringBuilder sb = new StringBuilder();
		for (int i = 0; i < 400; i++) {
			sb.append(i + "bla bla" + i + " ");
		}
		par.getContent().add(StyleFactory.textFormatting(sb.toString(), DocxStyleEnum.TEXT_NORMAL));

		P par1 = factory.createP();
		par1.getContent().add(StyleFactory.customizeLayoutSection(1));
		sb = new StringBuilder();
		for (int i = 0; i < 23; i++) {
			sb.append(i + "have a nice day ");
		}
		par1.getContent().add(StyleFactory.textFormatting(sb.toString(), DocxStyleEnum.TEXT_SUB_TITLE));

		wordMLPackage.getMainDocumentPart().getContent().add(par);
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());
		wordMLPackage.getMainDocumentPart().getContent().add(par1);
		wordMLPackage.save(f);
		log(f.getAbsolutePath(), true);
	}

	@Test
	public void completeReport() throws Exception {
		String newFileName = "05.docx";

		File f = new File(SRC_RESOURCES + newFileName);
		if (f.exists())
			f.delete();

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		StyleFactory.overrideDefaultStyle(wordMLPackage.getMainDocumentPart(), new CustomNormalStyle(), "normal");
		log("Override [normal] style", true);
		StyleFactory.overrideDefaultStyle(wordMLPackage.getMainDocumentPart(), new CustomTitleStyle(), "title");
		log("Override [Title] style", true);
		StyleFactory.overrideDefaultStyle(wordMLPackage.getMainDocumentPart(), new CustomHeadingStyle(), "heading1");
		log("Override [heading1] style", true);

		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.customizeLayoutSection(1));
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "My Firts Custom Report");
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());

		Br objBr = new Br();
		objBr.setType(STBrType.PAGE);
		wordMLPackage.getMainDocumentPart().addObject(objBr);

		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "People that work");
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());

		P p3 = factory.createP();
		p3.getContent().add(StyleFactory.customizeLayoutSection(1));
		p3.getContent().add(StyleFactory.textFormatting("Digital transformation (DX) is fundamentally changing procurement, allowing businesses to transform their decision making, which is enhancing their business outcomes significantly as we enter an increasingly digital economy. Digital transformation is an enterprisewide, board-level, and strategic reality for companies wishing to remain relevant or enhance their leadership position in the digitaleconomy. Digitally transformed businesses have a repeatable set of practices and disciplines used to leverage new business", new CustomNormalStyle()));
		p3.getContent().add(StyleFactory.textFormatting(" and markets in pursuit of business performance and growth ", DocxStyleEnum.TEXT_BOLD));
		p3.getContent().add(StyleFactory.textFormatting("Digital transformation (DX) is fundamentally changing procurement, allowing businesses to transform their decision making, which is enhancing their business outcomes significantly as we enter an increasingly digital economy. Digital transformation is an enterprisewide, board-level, and strategic reality for companies wishing to remain relevant or enhance their leadership position in the digitaleconomy. Digitally transformed businesses have a repeatable set of practices and disciplines used to leverage new business", new CustomNormalStyle()));
		wordMLPackage.getMainDocumentPart().getContent().add(p3);

		byte[] bytes = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "1.png");
		byte[] bytes1 = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "2.png");
		byte[] bytes2 = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "3.png");
		P image1 = StyleFactory.newImage(wordMLPackage, bytes, "caption", "caption", 1, 1, 500);
		P image2 = StyleFactory.newImage(wordMLPackage, bytes1, "caption", "caption", 1, 1, 500);
		P image3 = StyleFactory.newImage(wordMLPackage, bytes2, "caption", "caption", 1, 1, 500);

		log("create image", true);
		List<String> header = new ArrayList<>();
		header.add("Company Name");
		header.add("");
		header.add("description");
		header.add("Link");

		List<Object> cellValue = new ArrayList<>();
		cellValue.add("Andrea Radice");
		cellValue.add(image1);
		cellValue.add("The best");
		cellValue.add(StyleFactory.createHyperlink(wordMLPackage.getMainDocumentPart(), "http://radic.altervista.org", "My webSite"));
		cellValue.add("Micky Mouse");
		cellValue.add(image2);
		cellValue.add("SVP Research & Development at Disney");
		cellValue.add(StyleFactory.createHyperlink(wordMLPackage.getMainDocumentPart(), "http://www.google.it", ""));
		cellValue.add("Donal Duck");
		cellValue.add(image3);
		cellValue.add("Funny Duck");
		cellValue.add("");

		Tbl table = StyleFactory.createTableWithContent(DocxStyleEnum.TABLE_HEADER_CELL, DocxStyleEnum.TABLE_EVEN_CELL, DocxStyleEnum.TABLE_ODD_CELL, header, cellValue);
		table.getContent().add(StyleFactory.customizeLayoutSection(1));
		wordMLPackage.getMainDocumentPart().addObject(table);
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());
		log("created table", true);

		wordMLPackage.getMainDocumentPart().addObject(objBr);
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "World");
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());

		P par = factory.createP();
		par.getContent().add(StyleFactory.customizeLayoutSection(2));
		par.getContent().add(StyleFactory.textFormatting("The world is the planet Earth and all life upon it, including human civilization. In a philosophical context, the \"world\" is the whole of the physical Universe, or an ontological world (the \"world\" of an individual). In a theological context, the world is the material or the profane sphere, as opposed to the celestial, spiritual, transcendent or sacred. The \"end of the world\" refers to scenarios of the final end of human history, often in religious contexts.History of the world is commonly understood as spanning the major geopolitical developments of about five millennia, from the first civilizations to the present. In terms such as world religion, world language, world government, and world war, world suggests international or intercontinental scope without necessarily implying participation of the entire world.", DocxStyleEnum.TEXT_NORMAL));
		wordMLPackage.getMainDocumentPart().getContent().add(par);
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());
		P par1 = factory.createP();
		par1.getContent().add(StyleFactory.customizeLayoutSection(1));
		par1.getContent().add(StyleFactory.textFormatting("The Palestinian health ministry says at least five people have been killed and 350 wounded, many of them by Israeli gunfire. The Israeli military reported \"rioting\" at six places and said it was \"firing towards main instigators\". Palestinians have pitched five camps near the border for the protest, dubbed the Great March of Return.", DocxStyleEnum.TEXT_BOLD));
		wordMLPackage.getMainDocumentPart().getContent().add(par1);
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());

		byte[] bytesW = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "1.png");
		P imageW = StyleFactory.newImage(wordMLPackage, bytesW, "caption", "caption", 1, 1, 10000);
		wordMLPackage.getMainDocumentPart().getContent().add(imageW);

		wordMLPackage.getMainDocumentPart().addObject(objBr);
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "Europe");
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());
		byte[] bytesE = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "2.png");
		P imageE = StyleFactory.newImage(wordMLPackage, bytesE, "caption", "caption", 1, 1, 10000);
		wordMLPackage.getMainDocumentPart().getContent().add(imageE);
		wordMLPackage.getMainDocumentPart().addObject(StyleFactory.createSectionBreak());
		wordMLPackage.getMainDocumentPart().addParagraphOfText("Europe is a Swedish rock band formed in Upplands Väsby in 1979,[5] by vocalist Joey Tempest, guitarist John Norum, bass guitarist Peter Olsson, and drummer Tony Reno. They got a major breakthrough in Sweden in 1982 by winning the televised competition \"Rock-SM\" (Swedish Rock Championships) It was the first time this competition was held, and Europe became a larger success than the competition itself. Since their formation Europe has released eleven studio albums, three live albums, three compilations and nineteen music videos.");
		wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "America");

		byte[] bytesA = StyleFactory.getImageBytes(SRC_RESOURCES_IMAGE + "ok.jpg");
		P imageA = StyleFactory.newImage(wordMLPackage, bytesA, null, null, 1, 1, 10000);
		wordMLPackage.getMainDocumentPart().getContent().add(imageA);

		List<String> header1 = new ArrayList<>();
		header1.add("num");
		header1.add("Style");
		header1.add("Color");

		List<Object> cellValue1 = new ArrayList<>();
		Random rand = new Random();
		for (int i = 0; i < 30; i++) {
			cellValue1.add("" + i);
			cellValue1.add(randomFont());
			cellValue1.add(randomColor(rand));
		}

		Tbl table2 = StyleFactory.createTableWithContent(DocxStyleEnum.TABLE_NO_BORDER, DocxStyleEnum.TABLE_NO_BORDER, DocxStyleEnum.TABLE_NO_BORDER, header1, cellValue1);
		wordMLPackage.getMainDocumentPart().addObject(table2);

		wordMLPackage.save(f);
		log(f.getAbsolutePath(), true);
	}


	/* Utility method */
	private static String randomColor(Random r) {
		StringBuilder color = new StringBuilder(Integer.toHexString(r
				.nextInt(16777215)));
		while (color.length() < 6) {
			color.append("0");
		}
		return color.append("").reverse().toString();
	}

	private static String randomFont() {
		switch (RandomUtils.nextInt(0, 4)) {
		case 0:
			return "Verdana";
		case 1:
			return "Tahoma";
		case 2:
			return "Calibri";
		case 3:
			return "Courier new";
		default:
			return "Arial";
		}
	}

	private static void log(String text, boolean success) {
		System.out.println((success ? "[OK] - " : "[ERROR] - ") + text);
	}
}
