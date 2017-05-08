package BlockEditMode;

public interface BlockHistoryOperationHandler {
	
	public byte[] getVariantContentByName(String variantName);

	public boolean isListOnlyContent(byte[] variantContent);

	public byte[] addParagraphOnTheEnd(byte[] variantContent);

}
