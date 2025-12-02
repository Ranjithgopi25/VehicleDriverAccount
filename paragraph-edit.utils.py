import { ParagraphEdit } from '../models/message.model';

/**
 * Split content into paragraphs using double newlines
 * @param content - Content text to split
 * @returns Array of paragraph strings
 */
export function splitIntoParagraphs(content: string): string[] {
  if (!content || !content.trim()) {
    return [];
  }
  
  const paragraphs = content.split('\n\n').map(p => p.trim()).filter(p => p.length > 0);
  
  // If no double newlines found, treat entire content as single paragraph
  if (paragraphs.length === 0) {
    return [content.trim()];
  }
  
  return paragraphs;
}

/**
 * Create paragraph edits by comparing original and edited content
 * @param original - Original content text
 * @param edited - Edited content text
 * @param editorNames - Array of editor names to include in tags
 * @returns Array of ParagraphEdit objects
 */
export function createParagraphEditsFromComparison(
  original: string,
  edited: string,
  editorNames: string[]
): ParagraphEdit[] {
  if (!original || !original.trim()) {
    // If no original content, create edits from edited content only
    const editedParagraphs = splitIntoParagraphs(edited);
    return editedParagraphs.map((editedPara, i) => ({
      index: i,
      original: '',
      edited: editedPara,
      tags: editorNames.map(name => `${name} (Reviewed)`),
      approved: null
    }));
  }
  
  const originalParagraphs = splitIntoParagraphs(original);
  const editedParagraphs = splitIntoParagraphs(edited);
  
  const paragraphEdits: ParagraphEdit[] = [];
  const maxLen = Math.max(originalParagraphs.length, editedParagraphs.length);
  
  for (let i = 0; i < maxLen; i++) {
    // Ensure we always have a value, even if empty string
    const originalPara = (i < originalParagraphs.length && originalParagraphs[i]) ? originalParagraphs[i].trim() : '';
    const editedPara = (i < editedParagraphs.length && editedParagraphs[i]) ? editedParagraphs[i].trim() : '';
    const paragraphChanged = originalPara !== editedPara;
    
    // Always include ALL editors that were used
    const tags = editorNames.map(editorName => {
      if (paragraphChanged) {
        return `${editorName} (Editorial rule)`;
      } else {
        return `${editorName} (Reviewed)`;
      }
    });
    
    paragraphEdits.push({
      index: i,
      original: originalPara,
      edited: editedPara,
      tags: tags,
      approved: null
    });
  }
  
  return paragraphEdits;
}

/**
 * Check if all paragraphs have been decided (approved or declined)
 * @param paragraphEdits - Array of paragraph edits
 * @returns True if all paragraphs have been decided
 */
export function allParagraphsDecided(paragraphEdits: ParagraphEdit[]): boolean {
  return paragraphEdits.length > 0 && 
         paragraphEdits.every(p => p.approved !== null);
}

