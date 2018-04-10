package it.naturtalent.libreoffice;


public class DrawDocumentEvent 
{

	public static final String DRAWDOCUMENT_EVENT = "drawEvent/"; //$NON-NLS-N$
	
	// der Ladeprozess ist abgeschlossen (vor Speichern in OpenDocument Tabelle  @see OpenDesignAction)	
	public static final String DRAWDOCUMENT_EVENT_DOCUMENT_JUSTOPENED = DRAWDOCUMENT_EVENT+"drawDocumentJustOpened"; //$NON-NLS-N$
	
	// ein DrawDocument wurde geoffnet (nach dem Speichern in OpenDocument Tabelle @see OpenDesignAction)
	public static final String DRAWDOCUMENT_EVENT_DOCUMENT_OPEN = DRAWDOCUMENT_EVENT+"drawDocumentOpen"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_DOCUMENT_CLOSE = DRAWDOCUMENT_EVENT+"drawDocumentClose"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_DOCUMENT_OPEN_CANCEL = DRAWDOCUMENT_EVENT+"drawDocumentOpenCancel"; //$NON-NLS-N$
	
	//public static final String DRAWDOCUMENT_EVENT_DOCUMENT_ADDED = DRAWDOCUMENT_EVENT+"drawDocumentAdded"; //$NON-NLS-N$
	
	//public static final String DRAWDOCUMENT_EVENT_DOCUMENT_MODIFIED = DRAWDOCUMENT_EVENT+"drawDocumentModified"; //$NON-NLS-N$
	
	//public static final String DRAWDOCUMENT_EVENT_SCALEDENOMINATOR_CHANGED = DRAWDOCUMENT_EVENT+"drawDocumentScaleDenominator"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_SHAPE_PULLED = DRAWDOCUMENT_EVENT+"drawDocumentshapepulled"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_SHAPE_SELECTED = DRAWDOCUMENT_EVENT+"drawDocumentshapeselected"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_DOCUMENT_CHANGEREQUEST = DRAWDOCUMENT_EVENT+"drawDocumentchangerequest"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_DOCUMENT_ACTIVATE = DRAWDOCUMENT_EVENT+"drawDocumentActivate"; //$NON-NLS-N$
	
	public static final String DRAWDOCUMENT_EVENT_GLOBALMOUSEPRESSED = DRAWDOCUMENT_EVENT+"globalmousecklick"; //$NON-NLS-N$
	
	
	
	public static final String DRAWDOCUMENT_PAGECHANGE_PROPERTY = "pagechangeproperty"; //$NON-NLS-N$
}
