// src/services/AiService.ts
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointDocumentService } from './SharePointDocumentService';

interface AIResponse {
  success: boolean;
  message: string;
  response?: string;
  error?: string;
}

export interface SharePointDocument {
  id: string;
  name: string;
  url: string;
  contentType: string;
  createdDate: Date;
}

export class AiService {
  private context: WebPartContext;
  private apiKey: string;
  private apiVersion: string;
  private azureEndpoint: string;
  private model: string;

  constructor(context: WebPartContext) {
    this.context = context;
    
    // Direkt hårdkodade värden - enklast för utveckling 
    this.apiKey = process.env.AZURE_OPENAI_API_KEY || "";
    this.apiVersion = "2024-10-21";
    this.azureEndpoint = "https://openai-sales-project.openai.azure.com/";
    this.model = "gpt-4o";
  }

  /**
   * Gets AI-generated content based on user input and SharePoint context
   */
  public async getAIResponse(userPrompt: string, referenceDocuments: string[]): Promise<AIResponse> {
    try {
      // Generate a complete prompt that includes context from SharePoint
      const enhancedPrompt = await this.buildEnhancedPrompt(userPrompt, referenceDocuments);
      
      // Prepare the request payload
      const payload = {
        model: this.model,
        messages: [
          {
            role: "system", 
            content: `Du är en assistent som hjälper till att skapa professionellt innehåll.`
          },
          {
            role: "user",
            content: enhancedPrompt
          }
        ],
        temperature: 0.7,
        max_tokens: 4096
      };

      // Make the API call
      const response = await fetch(`${this.azureEndpoint}openai/deployments/${this.model}/chat/completions?api-version=${this.apiVersion}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": this.apiKey
        },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: { message: `HTTP error! Status: ${response.status}` }}));
        throw new Error(errorData.error?.message || `API request failed with status: ${response.status}`);
      }

      const data = await response.json();
      const aiContent = data.choices[0].message.content;
      console.log("AI-innehåll (första 200 tecken):", aiContent.substring(0, 200));
      console.log("AI-innehåll längd:", aiContent.length);
      console.log("Innehåller endast ASCII?", /^[\x00-\x7F]*$/.test(aiContent));
      console.log("AI-innehåll:", aiContent.substring(0, 200) + "...");
      return {
        success: true,
        message: "AI response generated successfully",
        response: data.choices[0].message.content
      };
    } catch (error) {
      console.error("Error calling AI API:", error);
      return {
        success: false,
        message: "Failed to generate AI response",
        error: error instanceof Error ? error.message : "Unknown error occurred"
      };
    }
  }

  /**
   * Fetches SharePoint document templates from the current site
   */
  public async getWordTemplates(): Promise<SharePointDocument[]> {
    try {
      // Get Word templates from the current site
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Dokument')/items?$filter=endswith(File/Name, '.docx')&$expand=File&$select=Id,File/Name,File/ServerRelativeUrl`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch templates: ${response.statusText}`);
      }

      const data = await response.json();
      const aiContent = data.choices[0].message.content;
      console.log("AI-innehåll:", aiContent.substring(0, 200) + "...");
      
      return data.value.map((item:any) => ({
        id: item.Id,
        name: item.File.Name,
        url: item.File.ServerRelativeUrl,
        contentType: 'Word Document',
        createdDate: new Date(item.Created)
      }));
    } catch (error) {
      console.error("Error fetching Word templates:", error);
      return [];
    }
  }

  /**
   * Fetches content from SharePoint pages to use as reference material
   */
  public async getSharePointPageContent(pageUrl: string): Promise<string> {
    try {
      // This is a simplified approach - in reality, you would need to
      // parse the HTML content to extract relevant text
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        pageUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch page content: ${response.statusText}`);
      }

      const data = await response.text();
      
      // Extract text content from HTML (simplified approach)
      const textContent = data.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
      
      return textContent;
    } catch (error) {
      console.error("Error fetching SharePoint page content:", error);
      return "";
    }
  }

  /**
   * Get content from the current SharePoint page where the web part is placed
   */

  // @ts-ignore - Needed for future use
  private async getCurrentPageContent(): Promise<string> {
    try {
      // Get the current page URL from context - använd serverRequestPath istället för page.title
      const currentPageUrl = this.context.pageContext.web.serverRelativeUrl;
      // Vi kan inte använda page.title, så vi extraherar sidnamnet från URL:en
      const urlParts = this.context.pageContext.site.serverRequestPath.split('/');
      const currentPageName = urlParts[urlParts.length - 1] || 'Current Page';
      
      console.log("Fetching content from current page:", currentPageUrl);
      
      // Use SharePoint REST API to get the page content
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${currentPageUrl}')/ListItemAllFields`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.warn(`Could not fetch current page content: ${response.statusText}`);
        return "Could not fetch current page content.";
      }

      const pageData = await response.json();
      
      // Try to extract page content - this depends on how your pages are structured
      // This is a simplified attempt to get the main content
      let pageContent = "";
      
      if (pageData.WikiField) {
        // For wiki pages
        pageContent = pageData.WikiField;
      } else if (pageData.CanvasContent1) {
        // For modern pages
        pageContent = pageData.CanvasContent1;
      } else {
        // General fallback
        pageContent = JSON.stringify(pageData);
      }
      
      // Clean up the content
      const textContent = pageContent.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
      
      return `Current page (${currentPageName}): ${textContent}`;
    } catch (error) {
      console.error("Error fetching current page content:", error);
      return "Error fetching current page content.";
    }
  }  

  /**
   * Scans the current SharePoint site for related documents
   */

  // @ts-ignore - Needed for future use
  private async scanSiteForRelatedDocuments(keywords: string[]): Promise<string> {
    try {
      // This is a simplified search functionality that looks for documents
      // containing any of the keywords in their title or content
      if (!keywords || keywords.length === 0) {
        return "";
      }
      
      // Create a search query based on keywords
      const searchTerms = keywords.join(" OR ");
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${encodeURIComponent(searchTerms)}'&selectproperties='Title,Path'&rowlimit=5`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.warn(`Search for related documents failed: ${response.statusText}`);
        return "";
      }

      const searchData = await response.json();
      
      // Process search results
      let relatedContent = "Related documents found on site:\n";
      
      if (searchData.PrimaryQueryResult &&
          searchData.PrimaryQueryResult.RelevantResults &&
          searchData.PrimaryQueryResult.RelevantResults.Table &&
          searchData.PrimaryQueryResult.RelevantResults.Table.Rows) {
            
        const results = searchData.PrimaryQueryResult.RelevantResults.Table.Rows;
        
        for (const row of results) {
          const cells = row.Cells;
          const title = cells.find((c:any) => c.Key === 'Title')?.Value || 'Unnamed Document';
          const path = cells.find((c:any) => c.Key === 'Path')?.Value || '';
          
          // Add to related content
          relatedContent += `- ${title} (${path})\n`;
          
          // Try to fetch and include some content from this document
          if (path.endsWith('.docx')) {
            try {
              const content = await this.getSharePointPageContent(path);
              relatedContent += `  Summary: ${content.substring(0, 200)}...\n`;
            } catch (error) {
              // Continue even if one document fails
              console.error(`Error fetching content for ${path}:`, error);
            }
          }
        }
      } else {
        relatedContent += "No directly related documents found.";
      }
      
      return relatedContent;
    } catch (error) {
      console.error("Error scanning site for related documents:", error);
      return "Error scanning site for related documents.";
    }
  }

  /**
   * Creates an enhanced prompt that includes context from the current SharePoint page 
   * and any explicitly referenced documents
   */
  private async buildEnhancedPrompt(userPrompt: string, referenceDocUrls: string[]): Promise<string> {
    // Get content from the current page
    //const currentPageContent = await this.getCurrentPageContent();
    
    // Extract key terms for searching related documents
    // @ts-ignore - Needed for future use
    const keyTerms = userPrompt.split(' ')
      .filter(word => word.length > 4)
      .map(word => word.toLowerCase())
      .slice(0, 5); // Take up to 5 key terms
    
    // Scan for related documents
    //const relatedDocContent = await this.scanSiteForRelatedDocuments(keyTerms);
    
    // Get content from explicitly referenced documents
    let referenceContent = "";
    for (const url of referenceDocUrls) {
      const content = await this.getSharePointPageContent(url);
      referenceContent += `\n\nContent from ${url}:\n${content}`;
    }

    // Build a rich prompt that includes all SharePoint context
    return `
    User request: ${userPrompt}

    CURRENT SHAREPOINT CONTEXT:

    EXPLICITLY REFERENCED MATERIALS:
    ${referenceContent ? referenceContent : "No documents explicitly selected for reference."}

    Please generate content that matches the style and format of our existing documents. 
    The content should be professional and ready to insert into a Word document.

    Generate a complete and professional response that fulfills the user's request while 
    incorporating relevant information from the provided SharePoint context and reference materials.
    `;
  } 

  /**
   * Create a new Word document in SharePoint with AI-generated content
   */
  /**
 * Create a new Word document in SharePoint with AI-generated content
 */
  public async createWordDocument(title: string, content: string, templateUrl?: string): Promise<string> {
    try {
      // Använd SharePointDocumentService istället för egen implementation
      const spDocService = new SharePointDocumentService(this.context);
      
      // Använd createWordDocumentWithGraph-metoden istället för createWordDocument
      return await spDocService.createWordDocument(
        content,
        title,
        "Shared Documents" // Använd "Shared Documents" istället för "Dokument"
      );
    } catch (error) {
      console.error("Error creating Word document:", error);  
      throw error;
    }
  }
}