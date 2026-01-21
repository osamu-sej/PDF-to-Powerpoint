import { GoogleGenAI } from "@google/genai";

const wait = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export const generateSlideNotes = async (
  base64Image: string,
  pageNumber: number,
  userApiKey?: string
): Promise<string> => {
  // Use user-provided key if available, otherwise fall back to environment variable
  const apiKey = userApiKey || process.env.API_KEY;
  
  if (!apiKey) {
    console.warn("API Key not found, skipping AI notes.");
    return "";
  }

  const ai = new GoogleGenAI({ apiKey });
  
  // Clean base64 string if it contains metadata prefix
  const cleanBase64 = base64Image.split(',')[1];
  
  const maxRetries = 3;
  let attempt = 0;

  while (attempt <= maxRetries) {
    try {
      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: {
          parts: [
            {
              inlineData: {
                mimeType: 'image/jpeg',
                data: cleanBase64,
              },
            },
            {
              text: `Analyze this slide (Page ${pageNumber}). 
              1. Provide a concise summary of the key points suitable for "Speaker Notes".
              2. If there are charts or visual data, briefly describe the trend.
              3. Keep the tone professional.
              4. Do not include markdown formatting like **bold** or headers, just plain text paragraphs.`
            },
          ],
        },
        config: {
          thinkingConfig: { thinkingBudget: 0 } // Disable thinking for faster response on simple tasks
        }
      });

      return response.text || "";
    } catch (error: any) {
      const isRateLimit = error.status === 429 || error.code === 429 || (error.message && error.message.includes('429'));
      const isServerOverloaded = error.status === 503 || error.code === 503;

      if (isRateLimit || isServerOverloaded) {
        attempt++;
        if (attempt > maxRetries) {
          console.error(`Failed to generate notes for page ${pageNumber} after ${maxRetries} retries:`, error);
          return "";
        }
        
        // Exponential backoff: 2s, 4s, 8s...
        const delayMs = Math.pow(2, attempt) * 1000; 
        console.warn(`Rate limit hit for page ${pageNumber}. Retrying in ${delayMs}ms... (Attempt ${attempt}/${maxRetries})`);
        await wait(delayMs);
        continue;
      }

      console.error(`Error generating notes for page ${pageNumber}:`, error);
      return ""; // Fail gracefully without breaking the app
    }
  }
  return "";
};