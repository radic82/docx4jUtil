package it.docx4j.utils;

import org.docx4j.diff.Differencer;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.CTRPrChange;
import org.docx4j.wml.CTSimpleField;
import org.docx4j.wml.CTTrackChange;
import org.docx4j.wml.DelText;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RunDel;
import org.docx4j.wml.RunIns;
import org.docx4j.wml.Text;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class Docx4JXmlWrapper {
	private final static Boolean SHOW_MESSAGE = true;

	/**
	 * Retrive Bookmark
	 *
	 * @param jaxbObj JAXBElement
	 * @return bookmark
	 */
	public static CTBookmark getBookmark(JAXBElement<?> jaxbObj) {
		if (!(jaxbObj.getValue() instanceof CTBookmark)) {
			return null;
		}
		return (CTBookmark) jaxbObj.getValue();
	}

	private static RunIns getRevisionIns(Object jaxbObj) {
		return (RunIns) jaxbObj;
	}

	private static RunDel getRevisionDel(Object jaxbObj) {
		return (RunDel) jaxbObj;
	}

	private static CTTrackChange getRevision(Object jaxbObj) {
		return (CTTrackChange) jaxbObj;
	}

	public static CTSimpleField createSimpleField(String instructions, String currentText) {
		CTSimpleField ctSimple = new CTSimpleField();
		ctSimple.setInstr(instructions);

		ObjectFactory factory = Context.getWmlObjectFactory();

		Text t = factory.createText();
		t.setValue(currentText);

		R run = factory.createR();
		run.getRunContent().add(t);

		ctSimple.getParagraphContent().add(run);

		return ctSimple;
	}

	public static void bookmarkRun(P p, R r, String name, int id) {
		// Find the index
		int index = p.getContent().indexOf(r);

		if (index < 0) {
			System.out.println("P does not contain R!");
			return;
		}

		ObjectFactory factory = Context.getWmlObjectFactory();
		BigInteger ID = BigInteger.valueOf(id);

		// Add bookmark end first
		CTMarkupRange mr = factory.createCTMarkupRange();
		mr.setId(ID);
		JAXBElement<CTMarkupRange> bmEnd = factory.createBodyBookmarkEnd(mr);
		p.getContent().add(index + 1, bmEnd);

		// Next, bookmark start
		CTBookmark bm = factory.createCTBookmark();
		bm.setId(ID);
		bm.setName(name);
		JAXBElement<CTBookmark> bmStart = factory.createBodyBookmarkStart(bm);
		p.getContent().add(index, bmStart);
	}

	public static Map<String, String> createRevision(R r, String changeType) {
		int changeNum = 1;
		Map<String, String> changesMap = new HashMap<>();
		Object objInternal = r.getContent();
		if (objInternal instanceof JAXBElement) {
			changesMap.put(changeType + changeNum++, "");
		} else if (objInternal instanceof ArrayList) {
			for (int j = 0; j < ((ArrayList) objInternal).size(); j++) {
				Object obj = ((ArrayList) objInternal).get(j);
				if (obj instanceof Text) {
					changesMap.put(changeType + changeNum++, ((Text) obj).getValue());
				} else if (obj instanceof DelText) {
					changesMap.put(changeType + changeNum++, ((DelText) obj).getValue());
				} else {
					JAXBElement jab = (JAXBElement) ((ArrayList) objInternal).get(j);
					if (jab.getValue() instanceof Text) {
						Text t = (Text) jab.getValue();
						changesMap.put(changeType + changeNum++, t.getValue());
					}
				}
			}
		}
		return changesMap;
	}

	public static void handleRels(Differencer pd, MainDocumentPart newMDP) throws Exception {
		RelationshipsPart rp = newMDP.getRelationshipsPart();

		// Since we are going to add rels appropriate to the docs being
		// compared, for neatness and to avoid duplication
		// (duplication of internal part names is fatal in Word,
		//  and export xslt makes images internal, though it does avoid duplicating
		//  a part ),
		// remove any existing rels which point to images
		List<Relationship> relsToRemove = new ArrayList<Relationship>();
		for (Relationship r : rp.getRelationships().getRelationship()) {
			//  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
			if (r.getType().equals(Namespaces.IMAGE)) {
				relsToRemove.add(r);
			}
		}
		for (Relationship r : relsToRemove) {
			rp.removeRelationship(r);
		}
		// Now add the rels we composed
		Map<Relationship, Part> newRels = pd.getComposedRels();
		for (Relationship nr : newRels.keySet()) {
			if (nr.getTargetMode() != null
					&& nr.getTargetMode().equals("External")) {
				newMDP.getRelationshipsPart().getRelationships().getRelationship().add(nr);
			} else {
				Part part = newRels.get(nr);
				if (part == null) {
					System.out.println("ERROR! Couldn't find part for rel " + nr.getId() + "  " + nr.getTargetMode());
				} else {
					if (part instanceof BinaryPart) { // ensure contents are loaded, before moving to new pkg
						((BinaryPart) part).getBuffer();
					}
					newMDP.addTargetPart(part, RelationshipsPart.AddPartBehaviour.RENAME_IF_NAME_EXISTS, nr.getId());
				}
			}
		}
	}

	public static Map<Long, Map<String, String>> readRevision(String docxFile) throws Exception {
		final String XPATH_TO_REVISION_INS = "//w:ins";
		final String XPATH_TO_REVISION_DEL = "//w:del";
		final String XPATH_TO_REVISION_FORMAT_CHANGE = "//w:rPrChange";

		Map<Long, Map<String, String>> returnMap = new TreeMap<>();

		org.docx4j.openpackaging.packages.WordprocessingMLPackage wordMLPackage = org.docx4j.openpackaging.packages.WordprocessingMLPackage.load(new File(docxFile));
		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

		List<?> revisionInsList = mainDocumentPart.getJAXBNodesViaXPath(XPATH_TO_REVISION_INS, true);

		for (Object obj : revisionInsList) {
			if (!(obj instanceof RunIns)) {
				if (SHOW_MESSAGE)
					System.out.println("readRevision: ins object not managed " + obj.getClass().getName());
				continue;
			}
			RunIns ins = Docx4JXmlWrapper.getRevisionIns(obj);
			if (SHOW_MESSAGE)
				System.out.print("ReadRevision: Insert id[" + ins.getId() + "] - ");

			Map<String, String> revisionChangesMap = new HashMap<>();
			List<Object> listR = ins.getCustomXmlOrSmartTagOrSdt();
			for (Object aListR : listR) {
				R r = ((R) aListR);
				revisionChangesMap = Docx4JXmlWrapper.createRevision(r, "INS");
			}
			if (ins.getId() != null) {
				returnMap.put(ins.getId().longValue(), revisionChangesMap);
				if (SHOW_MESSAGE)
					System.out.println(revisionChangesMap.values());
			} else {
				if (SHOW_MESSAGE)
					System.out.println("That's strange: RunIns hasn't an id");
			}
		}

		List<?> revisionDelList = mainDocumentPart.getJAXBNodesViaXPath(XPATH_TO_REVISION_DEL, false);

		for (Object obj : revisionDelList) {
			if (!(obj instanceof RunDel)) {
				if (SHOW_MESSAGE)
					System.out.println("readRevision: del object not managed " + obj.getClass().getName());
				continue;
			}
			RunDel del = Docx4JXmlWrapper.getRevisionDel(obj);
			if (SHOW_MESSAGE)
				System.out.print("ReadRevision: Delete id[" + del.getId() + "] - ");

			Map<String, String> revisionChangesMap = new HashMap<>();
			List<Object> listR = del.getCustomXmlOrSmartTagOrSdt();
			for (Object aListR : listR) {
				R r = ((R) aListR);
				revisionChangesMap = Docx4JXmlWrapper.createRevision(r, "DEL");
			}
			if (del.getId() != null) {
				returnMap.put(del.getId().longValue(), revisionChangesMap);
				if (SHOW_MESSAGE)
					System.out.println(revisionChangesMap.values());
			} else {
				if (SHOW_MESSAGE)
					System.out.println("That's strange: RunDel hasn't an id");
			}
		}

		List<?> revisionDChangeList = mainDocumentPart.getJAXBNodesViaXPath(XPATH_TO_REVISION_FORMAT_CHANGE, false);

		int changeNum = 1;
		for (Object obj : revisionDChangeList) {
			if (!(obj instanceof CTRPrChange)) {
				if (SHOW_MESSAGE)
					System.out.println("readRevision: change object not managed " + obj.getClass().getName());
				continue;
			}
			CTRPrChange change = (CTRPrChange) obj;
			if (SHOW_MESSAGE)
				System.out.print("ReadRevision: Modify Format id[" + change.getId() + "] - ");
			Map<String, String> revisionChangesMap = new HashMap<>();
			revisionChangesMap.put("OTHER_" + changeNum++, "");
			if (change.getId() != null) {
				returnMap.put(change.getId().longValue(), revisionChangesMap);
				if (SHOW_MESSAGE)
					System.out.println(revisionChangesMap.values());
			} else {
				if (SHOW_MESSAGE)
					System.out.println("That's strange: CTRPrChange hasn't an id");
			}
		}
		return returnMap;
	}

	private static void acceptRevision(CTTrackChange objRead) {
		Object parent = objRead.getParent();
		if (parent instanceof P) {
			((P) objRead.getParent()).getContent().remove(objRead);
			R r = new R();
			Text t = new Text();
			t.setValue(getText(objRead));
			r.getContent().add(t);
			((P) objRead.getParent()).getContent().add(r);
		} else if (parent instanceof CTSimpleField) {
			((CTSimpleField) objRead.getParent()).getContent().remove(objRead);
			R r = new R();
			Text t = new Text();
			t.setValue(getText(objRead));
			r.getContent().add(t);
			((CTSimpleField) objRead.getParent()).getContent().add(r);
		} else {
			System.out.println("WARNING parent.getClass() not managed");
		}
	}

	private static String getText(Object object) {
		StringBuilder textR = new StringBuilder();
		if (object instanceof RunIns) {
			for (Object aListR : ((RunIns) object).getCustomXmlOrSmartTagOrSdt()) {
				textR.append(internalGetText((R) aListR));
			}
		} else if (object instanceof RunDel) {
			for (Object aListR : ((RunDel) object).getCustomXmlOrSmartTagOrSdt()) {
				textR.append(internalGetText((R) aListR));
			}
		}
		return textR.toString();
	}

	private static void setText(Object object, String text) {
		if (object instanceof RunIns) {
			for (Object aListR : ((RunIns) object).getCustomXmlOrSmartTagOrSdt()) {
				internalSetText((R) aListR, text);
			}
		} else if (object instanceof RunDel) {
			for (Object aListR : ((RunDel) object).getCustomXmlOrSmartTagOrSdt()) {
				internalSetText((R) aListR, text);
			}
		}
	}

	private static String internalGetText(R r) {
		String returnString = "";
		Object objInternal = r.getContent();
		if (objInternal instanceof JAXBElement) {
			returnString = "";
		} else if (objInternal instanceof ArrayList) {
			for (int j = 0; j < ((ArrayList) objInternal).size(); j++) {
				Object obj = ((ArrayList) objInternal).get(j);
				if (obj instanceof Text) {
					returnString = ((Text) obj).getValue();
				} else if (obj instanceof DelText) {
					returnString = ((DelText) obj).getValue();
				} else {
					JAXBElement jab = (JAXBElement) ((ArrayList) objInternal).get(j);
					if (jab.getValue() instanceof Text) {
						Text t = (Text) jab.getValue();
						returnString = t.getValue();
					}
				}
			}
		}
		return returnString;
	}

	private static void internalSetText(R r, String text) {
		Object objInternal = r.getContent();
		if (objInternal instanceof ArrayList) {
			for (int j = 0; j < ((ArrayList) objInternal).size(); j++) {
				Object obj = ((ArrayList) objInternal).get(j);
				if (obj instanceof Text) {
					((Text) obj).setValue(text);
				} else if (obj instanceof DelText) {
					((DelText) obj).setValue(text);
				} else {
					JAXBElement jab = (JAXBElement) ((ArrayList) objInternal).get(j);
					if (jab.getValue() instanceof Text) {
						Text t = (Text) jab.getValue();
						t.setValue(text);
					}
				}
			}
		}
	}

	public static void insertAnchorRevision(String docxFile) throws Exception {
		final String XPATH_TO_REVISION_INS = "//w:ins";
		final String XPATH_TO_REVISION_DEL = "//w:del";
		final String XPATH_TO_REVISION_FORMAT_CHANGE = "//w:rPrChange";

		org.docx4j.openpackaging.packages.WordprocessingMLPackage wordMLPackage = org.docx4j.openpackaging.packages.WordprocessingMLPackage.load(new File(docxFile));
		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

		List<?> revisionInsList = mainDocumentPart.getJAXBNodesViaXPath(XPATH_TO_REVISION_INS, true);
		for (Object obj : revisionInsList) {
			if (obj instanceof RunIns) {
				RunIns ins = Docx4JXmlWrapper.getRevisionIns(obj);
				createAnchorNoParagraph(ins.getParent(), "#" + ins.getId() + "@");
			}
		}
		List<?> revisionDelList = mainDocumentPart.getJAXBNodesViaXPath(XPATH_TO_REVISION_DEL, false);

		for (Object obj : revisionDelList) {
			if (obj instanceof RunDel) {
				RunDel del = Docx4JXmlWrapper.getRevisionDel(obj);
				createAnchorNoParagraph(del.getParent(), "#" + del.getId() + "@");
			}
		}

		List<?> revisionDChangeList = mainDocumentPart.getJAXBNodesViaXPath(XPATH_TO_REVISION_FORMAT_CHANGE, false);
		for (Object obj : revisionDChangeList) {
			if (obj instanceof CTRPrChange) {
				CTRPrChange change = (CTRPrChange) obj;
				createAnchorNoParagraph(change.getParent(), "#" + change.getId() + "@");
			}
		}
		wordMLPackage.save(new File(docxFile));
	}

	public static void createAnchor(org.docx4j.openpackaging.packages.WordprocessingMLPackage wordMLPackage, String anchor) {
		try {
			ObjectFactory factory = Context.getWmlObjectFactory();
			P p1 = factory.createP();
			p1.getContent().add(createTextFormatting(anchor));
			wordMLPackage.getMainDocumentPart().getContent().add(p1);
		} catch (Exception e) {
			System.out.println("createAnchor exc[" + e.getMessage() + "] exc:" + e.getMessage());
		}
	}

	public static void createAnchorNoParagraph(Object obj, String anchor) {
		R r = createTextFormatting(anchor);
		if (obj instanceof P) {
			((P) obj).getContent().add(r);
		} else if (obj instanceof CTSimpleField) {
			((CTSimpleField) obj).getContent().add(r);
		} else {
			System.out.println("WARNING parent.getClass() not managed");
		}
	}

	private static R createTextFormatting(String anchor) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		R r = factory.createR();
		RPr rpr = factory.createRPr();
		org.docx4j.wml.Color color = new org.docx4j.wml.Color();
		color.setVal("ffff00"); //white
		rpr.setColor(color);
		HpsMeasure hpsMeasure = factory.createHpsMeasure();
		rpr.setSz(hpsMeasure);
		rpr.getSz().setVal(BigInteger.valueOf(12)); //no dimension
		r.setRPr(rpr);
		Text text = factory.createText();
		text.setValue(anchor);
		r.getContent().add(text);
		return r;
	}
}
