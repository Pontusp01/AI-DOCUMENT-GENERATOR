// src/services/WordDocumentGenerator.ts
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AiService } from './AiService';
import { SharePointDocumentService, SharePointDocument } from './SharePointDocumentService';

export interface WordGenerationRequest {
  userPrompt: string;
  selectedTemplateUrl?: string;
  documentName: string;
  referenceUrls: string[];
}

export interface WordGenerationResult {
  success: boolean;
  documentUrl?: string;
  error?: string;
}

export class WordDocumentGenerator {
  // @ts-ignore - Needed for future use
private context: WebPartContext;
// @ts-ignore - Needed for future use  
private aiService: AiService;
  private spDocService: SharePointDocumentService;
  
  constructor(context: WebPartContext) {
    this.context = context;
    this.aiService = new AiService(context);
    this.spDocService = new SharePointDocumentService(context);
  }

  /**
   * Generate a Word document based on AI content
   */
  public async generateDocument(request: WordGenerationRequest): Promise<WordGenerationResult> {
  try {
    // Kommentera bort anropet till AI-tjänsten
     const aiResponse = await this.aiService.getAIResponse(
       request.userPrompt,
       request.referenceUrls
     );
    
     if (!aiResponse.success || !aiResponse.response) {
       throw new Error(aiResponse.error || "Failed to generate AI content");
    }
    
    // 2. Create the Word document
    let documentUrl: string;
    
    if (request.selectedTemplateUrl) {
      // If a template was selected, use it
      documentUrl = await this.spDocService.createDocumentFromTemplate(
        request.selectedTemplateUrl,
        {
          // Define placeholder replacements based on the template structure
          'CONTENT': aiResponse.response,
          'TITLE': request.documentName,
          'DATE': new Date().toLocaleDateString()
        },
        request.documentName
      );
    } else {
      // Använd Graph API istället för vanlig createWordDocument
      documentUrl = await this.spDocService.createWordDocument(
        aiResponse.response,  // Använd det hårdkodade innehållet
        request.documentName,
        "Shared Documents"
      );
    }
    
    return {
      success: true,
      documentUrl
    };
  } catch (error) {
    console.error("Error generating Word document:", error);
    return {
      success: false,
      error: error instanceof Error ? error.message : "Unknown error generating document"
    };
  }
}

  /**
   * Gets available Word templates
   */
  public async getAvailableTemplates(): Promise<SharePointDocument[]> {
    try {
      return await this.spDocService.getWordDocumentsFromAllLibraries();
    } catch (error) {
      console.error("Error getting templates:", error);
      return [];
    }
  }

  /**
   * Gets available reference materials (pages and guides)
   */
  public async getReferenceMaterials(): Promise<SharePointDocument[]> {
    try {
      // Get site pages
      let materials: SharePointDocument[] = [];
      try {
        const sitePages = await this.spDocService.getSitePages();
        materials = [...materials, ...sitePages];
      } catch (error) {
        console.warn("Could not get site pages:", error);
      }

      // Get documents from all libraries
      try {
        const documents = await this.spDocService.getWordDocumentsFromAllLibraries();
        materials = [...materials, ...documents];
      } catch (error) {
        console.warn("Could not get documents:", error);
      }

      return materials;
    } catch (error) {
      console.error("Error getting reference materials:", error);
      return [];
    }
  }
}