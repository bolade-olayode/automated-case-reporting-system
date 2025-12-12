function uploadCasesToWordPress() {
  const WP_POSTS_URL = "";
  const WP_TAGS_URL = "";
  const USERNAME = "";
  const APP_PASSWORD = ""; 
  const FOLDER_ID = ""; 

  const CAT_JUDGMENTS = 23, CAT_SC = 890, CAT_CA = 20;
  const AUTH = "Basic " + Utilities.base64Encode(USERNAME + ":" + APP_PASSWORD);

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const originalFileId = file.getId();
    const fileName = file.getName();
    const originalMimeType = file.getMimeType(); 
    let docIdForParsing = null; 
    // More robust check for Google Docs MimeType variations
    let isOriginalGDoc = (originalMimeType === MimeType.GOOGLE_DOCS || originalMimeType === 'application/vnd.google-apps.document'); 


    Logger.log(`Processing file: ${fileName} (ID: ${originalFileId}, Type: ${originalMimeType})`);

    let gdoc; 

    if (isOriginalGDoc) {
      // If it's already a Google Doc, use it directly
      gdoc = file;
      docIdForParsing = originalFileId;
       Logger.log(`File ${fileName} is already a Google Doc. Using existing ID: ${docIdForParsing}`);
    } else {
      // If it's not a Google Doc (e.g., PDF), convert it
      gdoc = toGoogleDoc(file);
      if (!gdoc) {
          Logger.log(`Skipping file: ${fileName} - Could not convert to Google Doc.`);
          continue; // Skip to next file if conversion failed
      }
      docIdForParsing = gdoc.getId(); // Get the ID of the *newly created* Google Doc
       Logger.log(`Converted ${fileName} to new Google Doc ID: ${docIdForParsing}`);
    }

    // --- Permissions ---
    // Make sure the file being embedded (original) and parsed (gdoc) are viewable
    try {
      DriveApp.getFileById(originalFileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      // Only set sharing on converted doc if it's different and exists
      if (docIdForParsing && docIdForParsing !== originalFileId) { 
         DriveApp.getFileById(docIdForParsing).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
      Logger.log(`Set sharing for relevant docs related to ${fileName}`);
    } catch(e) {
      Logger.log(`Could not set file sharing permissions for ${fileName}: ${e}`);
    }
    
    // --- Parsing ---
    let text = "";
    if (!docIdForParsing) {
       Logger.log(`Skipping parsing for ${fileName} as docIdForParsing is null.`);
       continue;
    }
    try {
      const doc = DocumentApp.openById(docIdForParsing);
      const body = doc.getBody();
      if (body) {
        text = body.getText().replace(/\r/g, "");
         if (!text.trim()) {
           Logger.log(`Google Doc ${docIdForParsing} (from ${fileName}) is empty. Skipping.`);
           if (!isOriginalGDoc) { // Clean up empty converted doc
              try { DriveApp.getFileById(docIdForParsing).setTrashed(true); Logger.log(`Trashed empty converted doc ${docIdForParsing}`); } catch(err){}
           }
           continue;
         }
      } else {
         Logger.log(`Could not get body for Google Doc ${docIdForParsing} (from ${fileName}). Skipping.`);
          if (!isOriginalGDoc) { // Clean up failed converted doc
              try { DriveApp.getFileById(docIdForParsing).setTrashed(true); Logger.log(`Trashed failed converted doc ${docIdForParsing}`); } catch(err){}
           }
         continue;
      }
    } catch (e) {
       Logger.log(`Error reading text from Google Doc ${docIdForParsing} (from ${fileName}): ${e}`);
        if (!isOriginalGDoc) { // Clean up failed converted doc
              try { DriveApp.getFileById(docIdForParsing).setTrashed(true); Logger.log(`Trashed failed converted doc ${docIdForParsing} after read error.`); } catch(err){}
        }
       continue; // Skip this file if text cannot be read
    }

    const parsed = parseCase(text);
    if (!parsed) {
         Logger.log(`Error parsing content for file ${fileName}. Skipping.`);
          if (!isOriginalGDoc) { // Clean up failed converted doc
              try { DriveApp.getFileById(docIdForParsing).setTrashed(true); Logger.log(`Trashed failed converted doc ${docIdForParsing} after parse error.`); } catch(err){}
          }
         continue;
    }
    Logger.log(`Parsed content for ${fileName}. Ratio length: ${parsed.ratioHtml ? parsed.ratioHtml.length : 0}, Tags found: [${parsed.subjectMatters.join(', ')}]`);


    // --- Categories & Title ---
    let title = gdoc.getName().replace(/\.(docx?|pdf)$/i, '').trim(); 
    const categories = [CAT_JUDGMENTS];
    if (/-SC-/i.test(fileName)) categories.push(CAT_SC); 
    if (/-CA-/i.test(fileName)) categories.push(CAT_CA);

    // --- Tags ---
    const cleanTags = sanitizeSubjectMatters(parsed.subjectMatters);
    const tagIds = getOrCreateTagIds(cleanTags, WP_TAGS_URL, AUTH);
    Logger.log(`Tag IDs selected for ${fileName}: [${tagIds.join(', ')}]`);


    // --- HTML Building & Posting ---
    const html = buildPostHtml(parsed, originalFileId); 
    const payload = { title, content: html, status: "publish", categories, tags: tagIds };

    let postSuccess = false; 
    try {
        Logger.log(`Attempting to post "${title}" to WordPress...`);
        const response = UrlFetchApp.fetch(WP_POSTS_URL, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            headers: { "Authorization": AUTH },
            muteHttpExceptions: true 
        });
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
        if (responseCode >= 200 && responseCode < 300) {
            Logger.log(`Successfully posted "${title}" to WordPress. Response Code: ${responseCode}`);
            postSuccess = true; 
        } else {
            Logger.log(`Error posting "${title}" to WordPress. Status: ${responseCode}, Response: ${responseBody}`);
        }
    } catch (e) {
        Logger.log(`Network or severe error posting "${title}": ${e}`);
    }
    
    // --- Clean up Converted Doc ---
    if (postSuccess && !isOriginalGDoc && docIdForParsing) {
         try {
           DriveApp.getFileById(docIdForParsing).setTrashed(true);
           Logger.log(`Trashed CONVERTED Google Doc: ${docIdForParsing} for original file ${fileName}`);
         } catch (e) {
           Logger.log(`Failed to trash converted Google Doc ${docIdForParsing}: ${e}`);
         }
    } else if (!postSuccess && !isOriginalGDoc && docIdForParsing) {
         Logger.log(`WordPress post failed for ${fileName}. Leaving converted Google Doc ${docIdForParsing} for review.`);
    }

     Utilities.sleep(1500); 
  } // End while loop
  Logger.log("Finished processing all files in the folder.");
}

function toGoogleDoc(file) {
  const fileName = file.getName();
  const mimeType = file.getMimeType();
  // More robust check for Google Docs MimeType variations
  if (mimeType === MimeType.GOOGLE_DOCS || mimeType === 'application/vnd.google-apps.document') {
    Logger.log(`File ${fileName} is already a Google Doc (MimeType: ${mimeType}).`);
    return file; 
  }

  const convertibleTypes = [MimeType.PDF, MimeType.MICROSOFT_WORD, MimeType.RTF, MimeType.OPENDOCUMENT_TEXT];
  if (!convertibleTypes.includes(mimeType) && !fileName.toLowerCase().endsWith('.docx')) { 
      Logger.log(`Skipping conversion for ${fileName}. Unsupported MimeType for conversion: ${mimeType}`);
      return null;
  }

  Logger.log(`Converting ${fileName} (MimeType: ${mimeType}) to Google Doc...`);
  try {
    const blob = file.getBlob();
    const resource = { title: fileName, mimeType: MimeType.GOOGLE_DOCS };
    // Use convert: true; let Drive handle OCR implicitly when needed.
    const convertedFile = Drive.Files.insert(resource, blob, { convert: true }); 
    
    if (!convertedFile || !convertedFile.id) {
       Logger.log(`Conversion API call did not return a valid file ID for ${fileName}.`);
       return null;
    }
    Logger.log(`Successfully converted ${fileName} to Doc ID: ${convertedFile.id}`);
    Utilities.sleep(3000); // Increased delay after conversion
    return DriveApp.getFileById(convertedFile.id); 
  } catch (e) {
    Logger.log(`Error converting file ${fileName}: ${e.message} \n Stack: ${e.stack}`);
    if (e.message.includes("limit")) {
       Logger.log("Potential rate limit hit during conversion.");
       Utilities.sleep(6000); 
    }
    return null; 
  }
}


function parseCase(text) {
  try {
    const trimmedText = text.trim();
    if (!trimmedText) {
        Logger.log("parseCase: Input text is empty.");
        return null;
    }

    // --- Title Page Extraction ---
    const titleEndRegexes = [ /^\s*CASE\s+SYNOPSIS\b/mi, /^\s*INTRODUCTION\b/mi, /^\s*ISSUES?\b/mi, /^\s*RATIO\s+DECIDENDI\b/mi ];
    const titleEndIndex = findFirstMatchIndex(trimmedText, titleEndRegexes);
    const titlePageRaw = (titleEndIndex !== -1 ? trimmedText.substring(0, titleEndIndex) : trimmedText).trim(); 

    // --- Subject Matter Extraction ---
    const subjStartMatch = /Subject\s*Matter(?:\(s\))?\s*[:\s]*/i.exec(trimmedText);
    let subjectMatters = [];
    if (subjStartMatch) {
        const startIndex = subjStartMatch.index + subjStartMatch[0].length;
        const textAfterSubjStart = trimmedText.substring(startIndex);
        const subjEndRegexes = [ /^\s*Final\s*Order\b/mi, /^\s*CASE\s+SYNOPSIS\b/mi, /^\s*INTRODUCTION\b/mi, /^\s*ISSUES?\b/mi, /^\s*RATIO\s+DECIDENDI\b/mi ];
        const endIndex = findFirstMatchIndex(textAfterSubjStart, subjEndRegexes);
        const absoluteEndIndex = (endIndex !== -1) ? startIndex + endIndex : trimmedText.length;
        const subjRaw = trimmedText.substring(startIndex, absoluteEndIndex).trim();
        subjectMatters = parseSubjectMatters(subjRaw);
    }

    // --- Ratio Decidendi Extraction ---
    const ratioRaw = extractBetweenRegex( trimmedText, /\bRATIO\s+DECIDENDI\b/mi, /\b(?:FULL|MAIN)\s+JUDGMENTS?\b/mi );

    if (!titlePageRaw && !ratioRaw && subjectMatters.length === 0) {
        Logger.log("Parsing failed: Could not extract any required sections.");
        return null; 
    }

    return {
        titlePageHtml: styleTitlePage(titlePageRaw),
        ratioHtml: formatRatio(ratioRaw),
        subjectMatters: subjectMatters
    };
  } catch (e) {
      Logger.log(`Error during parsing: ${e.message} \n Stack: ${e.stack}`);
      return null; 
  }
}

function findFirstMatchIndex(text, regexList) {
    let firstIndex = -1;
    if (!text || !regexList || regexList.length === 0) return firstIndex;
    regexList.forEach(regex => {
        const match = regex.exec(text);
        if (match) {
            if (firstIndex === -1 || match.index < firstIndex) {
                firstIndex = match.index; 
            }
        }
    });
    return firstIndex;
}

function extractBetweenRegex(text, startRegex, endRegex) {
  if (!text) return "";
  const startMatch = startRegex.exec(text);
  if (!startMatch) return ""; 
  const startIndex = startMatch.index + startMatch[0].length;
  const textAfterStart = text.substring(startIndex);
  const endMatch = endRegex.exec(textAfterStart);
  if (!endMatch) return ""; // Return empty if end not found
  const endIndex = endMatch.index;
  const extracted = textAfterStart.substring(0, endIndex).trim();
  return extracted;
}


function styleTitlePage(raw) {
  if (!raw) return "";
  const BLUE = "#050298";
  const lines = raw.split("\n").map(s => s.trim()).filter(line => line && line.length > 1);
  let html = `<div style="text-align:center;">\n`;
  let foundCourtLine = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    // Section Headers
    if (/^Justices?\s*:?\s*$/i.test(line) || /^Subject\s*Matter(?:\(s\))?\s*:?\s*$/i.test(line) || /^Final\s*Order\s*:?\s*$/i.test(line) || /^Decision\s*:?\s*$/i.test(line)) { 
      html += `<div style="margin-top:20px;"></div>`; 
      html += `<p style="margin:14px 0 6px 0;"><span style="color:${BLUE}; font-weight:bold;">${escapeHtml(line)}</span></p>\n`;
      if (/^Subject\s*Matter/i.test(line)) foundCourtLine = true; 
    }
    // Court Name 
    else if (/^(Supreme Court|Court of Appeal)\b/i.test(line) && !foundCourtLine) {
      html += `<div style="margin-top:20px;"></div>`; 
      html += `<p style="margin:0; font-size:16px;"><span style="color:${BLUE}; font-weight:bold;">${escapeHtml(line)}</span></p>\n`;
      foundCourtLine = true; 
    }
    // Case Title, Citation (before Court Name)
    else if (!foundCourtLine) {
      html += `<p style="margin:0; font-size:16px;"><span style="color:${BLUE}; font-weight:bold;">${escapeHtml(line)}</span></p>\n`;
    }
    // Content under sections (Date, Justice names, Subject content, Order/Decision content)
    else {
      html += `<p style="margin:0 0 4px 0;">${escapeHtml(line)}</p>\n`; // Normal text
    }
  }
  return html + `</div>`;
}


function formatRatio(raw) {
  if (!raw) return ""; 
  const BLUE = "#050298";
  const lines = raw.split("\n").map(s => s.trim()).filter(line => line && line.length > 2);
  if (lines.length === 0) return "";

  let html = "";
  for (const line of lines) {
    // Check if the line is a subheading
    if (isRatioSubheading(line)) {
       // Apply bold style ONLY to subheadings
      html += `<p style="margin:12px 0 6px 0;"><span style="color:${BLUE}; font-weight:bold;">${escapeHtml(line)}</span></p>\n`;
    } else {
      // Apply blue color but NO bold style to regular text
      html += `<p style="margin:0 0 8px 0; color:${BLUE};">${escapeHtml(line)}</p>\n`; // Removed font-weight:bold
    }
  }
  return html;
}



function isRatioSubheading(line) {
  if (!line) return false;
  line = line.trim(); 
  // Rule 1: Ends with colon OR contains ALL CAPS before a dash/hyphen
  if (/:$/.test(line)) return true; 
  // Requires ALL CAPS before the dash, and the part before dash must be > 3 chars
  const dashMatch = line.match(/^([A-Z0-9\s\(\)/.,&-]+?)\s[-–—]\s(.+)$/); 
  if (dashMatch && dashMatch[1].trim().length > 3 && dashMatch[1].trim() === dashMatch[1].trim().toUpperCase()) {
      return true; 
  }

  // Rule 2: Mostly uppercase heuristic (tuned) - check for quotes and parentheses
  const letters = (line.match(/[A-Za-z]/g) || []).length;
  if (letters < 4) return false; 
  const uppers = (line.match(/[A-Z]/g) || []).length;
  // Requires >= 4 letters, > 80% uppercase, length < 250, does NOT contain quotes or parentheses
  if (letters >= 4 && (uppers / letters) > 0.80 && line.length < 250 && !line.includes('"') && !line.includes("'") && !line.includes('(') && !line.includes(')')) {
     return true; 
  }

  // Rule 3: Specific keywords (case-insensitive)
  if (/\b(OFFENCE|DEFENCE|EVIDENCE|JURISDICTION|APPEAL|MEANING OF|SCOPE OF|PRINCIPLE OF|BURDEN OF PROOF|STANDARD OF PROOF|ADMISSIBILITY|NEGLIGENCE|HOMICIDE|CONTRACT|ESTOPPEL|PRACTICE AND PROCEDURE|INTERPRETATION OF|CONSTITUTIONAL LAW|ACTION|COURT|JUDGMENT|ORDER|DAMAGES|LIMITATION LAW)\b/i.test(line)) {
      // Added check: Make sure it doesn't start with quote marks, which would indicate content not heading
      if (!/^\s*["“']/.test(line)) {
          return true;
      }
  }

  return false; 
}


function escapeHtml(s){
  if (!s || typeof s !== 'string') return '';
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

function parseSubjectMatters(subjRaw) {
  if (!subjRaw) return [];
  let s = subjRaw.replace(/\r/g, " ").replace(/\u2022|\*|\•/g, "\n").replace(/\s*;\s*/g, "\n").replace(/\s*,\s*(?=\d+\.?\s*)/g, "\n").replace(/(?<=\S)\s+(?=\d+\.?\s*)/g, "\n").replace(/(?<=\d)\s*\.\s*/g, ". ").replace(/^\s*\d+\s*[\)\.:\-]\s*/gm, "");
  const lines = s.split(/\n+/).map(x => x.trim()).filter(x => x && x.length > 5); 
  const uniqueLines = [...new Map(lines.map(item => [item.toLowerCase(), item])).values()];
  return uniqueLines;
}

function sanitizeSubjectMatters(items) {
  if (!items || !Array.isArray(items)) return [];
  const seen = new Set();
  const out = [];
  for (let raw of items) {
     if (typeof raw !== 'string') continue; 
    let p = raw.replace(/^\d+\s*[\)\.:\-]\s*/,"").replace(/[\.,;:]+$/,"").replace(/[()]/g, "").replace(/\s{2,}/g," ").replace(/^-\s*/, "").trim();
    if (!p || p.length < 5 || p.length > 100) continue; 
    const key = p.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      out.push(p); 
    }
    if (out.length >= 5) break; 
  }
  return out;
}

function getOrCreateTagIds(names, url, auth) {
  if (!names || !names.length) return [];
  const ids = [];
  const MAX_RETRIES_SEARCH = 2; 
  const MAX_RETRIES_CREATE = 1; 

  for (const n of names) {
    if (!n || typeof n !== 'string' || !n.trim()) continue; 
    const tagName = n.trim(); 
    let foundId = null;
    let searchAttempt = 0;
    let termExistsAfterFailedCreate = false;

    // Search Loop
    while (searchAttempt < MAX_RETRIES_SEARCH && foundId === null) {
        searchAttempt++;
        try {
            const searchUrl = `${url}?search=${encodeURIComponent(tagName)}&per_page=10&context=edit`; 
            const r1 = UrlFetchApp.fetch(searchUrl, { headers: { Authorization: auth }, muteHttpExceptions: true, validateHttpsCertificates: false });
            if (r1.getResponseCode() === 200) {
                const foundTags = JSON.parse(r1.getContentText());
                if (foundTags && foundTags.length > 0) {
                    const exactMatch = foundTags.find(t => t.name.trim().toLowerCase() === tagName.toLowerCase());
                    if (exactMatch) { foundId = exactMatch.id; break; } 
                }
            } else { Logger.log(`Error searching tag "${tagName}". Status: ${r1.getResponseCode()}`); break; }
        } catch (e) { Logger.log(`Exception during tag search "${tagName}": ${e}`); break; }
        if (foundId === null && searchAttempt < MAX_RETRIES_SEARCH) Utilities.sleep(500); 
    } 

    // Creation Logic
    if (foundId === null) {
        try {
            const createPayload = JSON.stringify({ name: tagName });
            const r2 = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: createPayload, headers: { Authorization: auth }, muteHttpExceptions: true, validateHttpsCertificates: false });
            if (r2.getResponseCode() === 201) { 
                const created = JSON.parse(r2.getContentText());
                if (created && created.id) foundId = created.id;
            } else if (r2.getResponseCode() === 400 && r2.getContentText().includes("term_exists")) {
                termExistsAfterFailedCreate = true; 
            } else { Logger.log(`Error creating tag "${tagName}". Status: ${r2.getResponseCode()}`); }
        } catch (e) { Logger.log(`Exception during tag creation "${tagName}": ${e}`); }
    }


     if (foundId === null && termExistsAfterFailedCreate) {
         Utilities.sleep(500); 
         try {
            const finalSearchUrl = `${url}?search=${encodeURIComponent(tagName)}&per_page=5&context=edit`;
            const r3 = UrlFetchApp.fetch(finalSearchUrl, { headers: { Authorization: auth }, muteHttpExceptions: true });
            if (r3.getResponseCode() === 200) {
                const finalFound = JSON.parse(r3.getContentText());
                const finalExactMatch = finalFound.find(t => t.name.trim().toLowerCase() === tagName.toLowerCase());
                if (finalExactMatch) foundId = finalExactMatch.id;
            }
         } catch (e) { Logger.log(`Exception during final tag search "${tagName}": ${e}`); }
     }

    if (foundId !== null) ids.push(foundId);
    else Logger.log(`---> FAILURE: Tag "${tagName}" not processed.`);

    Utilities.sleep(300); 
  } 
  return ids;
}



function buildPostHtml(parsed, fileId) {
  const BLUE = "#050298";
  const previewUrl = `https://drive.google.com/file/d/${fileId}/preview`; 
const titleHtml = (parsed && parsed.titlePageHtml) ? parsed.titlePageHtml : "<p><i>[Title page section could not be extracted.]</i></p>";
  const ratioHtmlContent = (parsed && parsed.ratioHtml) ? parsed.ratioHtml : "<p><i>[Ratio Decidendi section could not be extracted.]</i></p>"; 
  // Construct HTML with left-aligned Ratio heading
  let html = `
    ${titleHtml}
    
    <h3 style="color:${BLUE}; font-weight:bold; margin-top: 30px; text-align:left;">RATIO DECIDENDI</h3> 
    ${ratioHtmlContent}
    
    <div style="width:100%; max-width:100%; height:1200px; overflow:hidden; margin-top: 30px; border: 1px solid #ccc; position: relative;">
      <p style="text-align:center; font-weight:bold; color:${BLUE}; margin-bottom: 10px;">FULL JUDGMENT VIEW</p>
      <iframe 
        src="${previewUrl}" 
        width="100%" 
        height="100%" 
        frameborder="0" 
        scrolling="yes" 
        style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; border: none;"
        allow="fullscreen">
      </iframe>
    </div>
  `;
  return html;
}

