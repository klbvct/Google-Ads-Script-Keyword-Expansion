// Google Ads Script: Keyword Expansion Script - Campaign Level Check with OpenAI Classification
// This script fetches search terms, checks if they're already added as keywords
// in ANY ad group within the same campaign, applies performance thresholds,
// optionally classifies search terms using OpenAI, and exports qualifying terms to a Google Sheet
// Written by Matinique Roelse from Adcrease. Senior-only Google Ads agency.
// Linkedin: https://www.linkedin.com/in/matiniqueroelse/
// Website: https://www.adcrease.nl

// ===== CONFIGURATION =====
// IMPORTANT: If using OpenAI classification, please make a copy of this template sheet:
// https://docs.google.com/spreadsheets/d/16zYOrhS0MwSQu66kk2tcE6xUoTcG34_OFzrRwLy-Ff4/edit?gid=0#gid=0
// The template contains the required tabs and named ranges for the OpenAI API key.
const SHEET_URL = ''; // Leave empty to create a new spreadsheet
const TAB = 'Keyword Opportunities'; //Tab name

// ===== OPENAI CONFIGURATION =====
const USE_OPENAI_CLASSIFICATION = true; // Set to false to skip AI classification
const OPENAI_API_KEY_NAMED_RANGE = 'openaiapikey'; // Named range containing OpenAI API key
const WEBSITE_URL = ''; // Your website URL for content analysis (optional)

// OpenAI settings
const OPENAI_MODEL = 'gpt-3.5-turbo'; // Use gpt-4 for better accuracy (higher cost)
const BATCH_SIZE = 10; // Number of search terms to process in each API call
const MAX_TOTAL_COST = 5.00; // Maximum total cost in USD

// OpenAI pricing (as of 2024) - update these rates as needed
const OPENAI_PRICING = {
  "gpt-3.5-turbo": {
    input: 0.0015,  // $0.0015 per 1K input tokens
    output: 0.002   // $0.002 per 1K output tokens
  },
  "gpt-4": {
    input: 0.03,    // $0.03 per 1K input tokens
    output: 0.06    // $0.06 per 1K output tokens
  }
};

// Performance thresholds - modify these as needed
const MIN_COST = 5; // Minimum cost threshold (in currency units)
const MIN_CLICKS = 5; // Minimum clicks threshold
const MIN_CONVERSIONS = 1; // Minimum conversions threshold
const LOOKBACK_DAYS = 2; // Days to exclude from the end of the date range

// ===== DATE RANGE CONFIGURATION =====
// Choose ONE of the following date range options:

// Option 1: Use automatic date range (default 30 days with lookback)
const USE_AUTO_DATE_RANGE = true; // Set to false to use manual dates below
const NUM_DAYS = 90; // Total days to analyze (excluding lookback)

// Option 2: Use manual date range (set USE_AUTO_DATE_RANGE to false)
const MANUAL_START_DATE = '2025-01-01'; // Format: YYYY-MM-DD
const MANUAL_END_DATE = '2025-01-31';   // Format: YYYY-MM-DD

// Campaign filter - leave empty to include all enabled search campaigns
const CAMPAIGN_FILTER = ''; // e.g., 'brand' to only include campaigns containing 'brand'

// Campaign exclusion filter - leave empty to include all enabled search campaigns
const CAMPAIGN_EXCLUSION_FILTER = ''; // e.g., 'test' to exclude campaigns containing 'test'

// Target configuration - modify these values as needed
const TARGET_TYPE = 'ROAS'; // 'CPA' or 'ROAS'
const TARGET_VALUE = 2; // The target value for CPA or ROAS

// Global cost tracking
let totalCost = 0;
let totalInputTokens = 0;
let totalOutputTokens = 0;
let apiCallCount = 0;

function main() {
    try {
        // Reset cost tracking
        resetCostTracking();
        
        // Log target configuration
        Logger.log(`Using target: ${TARGET_TYPE} ${TARGET_TYPE === 'CPA' ? '<=' : '>='} ${TARGET_VALUE}`);
        
        // Get date range based on configuration
        const dateRange = getDateRange();
        
        // Build the search term query
        const searchTermQuery = buildSearchTermQuery(dateRange);
        
        Logger.log(`Executing search term query with date range: ${dateRange}`);
        Logger.log(`Campaign filter: ${CAMPAIGN_FILTER || 'None'}`);
        Logger.log(`Campaign exclusion filter: ${CAMPAIGN_EXCLUSION_FILTER || 'None'}`);
        Logger.log(`Target: ${TARGET_TYPE} ${TARGET_TYPE === 'CPA' ? '<=' : '>='} ${TARGET_VALUE}`);
        Logger.log(`OpenAI Classification: ${USE_OPENAI_CLASSIFICATION ? 'ENABLED' : 'DISABLED'}`);
        
        // Execute the search term query
        const searchTermRows = AdsApp.search(searchTermQuery);
        
        // Process the data with campaign-level keyword checking
        let data = processSearchTermsWithCampaignCheck(searchTermRows);
        
        // Apply OpenAI classification if enabled
        if (USE_OPENAI_CLASSIFICATION && data.length > 0) {
            data = applyOpenAIClassification(data);
        }
        
        // Export to spreadsheet
        exportToSheet(data);
        
        // Log cost summary
        logCostSummary();
        
        Logger.log(`Script completed successfully. Found ${data.length} qualifying search terms.`);
        
    } catch (e) {
        Logger.log(`Error in main function: ${e}`);
        throw e;
    }
}

function resetCostTracking() {
    totalCost = 0;
    totalInputTokens = 0;
    totalOutputTokens = 0;
    apiCallCount = 0;
}



function getDateRange() {
    if (USE_AUTO_DATE_RANGE) {
        // Use automatic date range with lookback
        return getDateRangeWithLookback(NUM_DAYS, LOOKBACK_DAYS);
    } else {
        // Use manual date range
        return getManualDateRange(MANUAL_START_DATE, MANUAL_END_DATE);
    }
}

function getDateRangeWithLookback(totalDays, lookbackDays) {
    const endDate = new Date();
    endDate.setDate(endDate.getDate() - lookbackDays); // Exclude lookback days
    
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - totalDays);
    
    const format = date => Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
    return `segments.date BETWEEN "${format(startDate)}" AND "${format(endDate)}"`;
}

function getManualDateRange(startDateStr, endDateStr) {
    try {
        // Parse the date strings
        const startDate = new Date(startDateStr);
        const endDate = new Date(endDateStr);
        
        // Validate dates
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            throw new Error('Invalid date format. Please use YYYY-MM-DD format.');
        }
        
        if (startDate > endDate) {
            throw new Error('Start date cannot be after end date.');
        }
        
        // Format dates for GAQL
        const format = date => Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
        return `segments.date BETWEEN "${format(startDate)}" AND "${format(endDate)}"`;
        
    } catch (e) {
        Logger.log(`Error with manual date range: ${e.message}`);
        Logger.log('Falling back to automatic date range...');
        return getDateRangeWithLookback(NUM_DAYS, LOOKBACK_DAYS);
    }
}

function buildSearchTermQuery(dateRange) {
    let query = `
SELECT 
    search_term_view.search_term,
    search_term_view.status,
    campaign.id,
    campaign.name,
    ad_group.id,
    ad_group.name,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros,
    metrics.conversions,
    metrics.conversions_value
FROM search_term_view
WHERE ${dateRange}
AND campaign.advertising_channel_type = "SEARCH"
AND campaign.status = "ENABLED"`;
    
    // Add campaign filter if specified
    if (CAMPAIGN_FILTER && CAMPAIGN_FILTER.trim() !== '') {
        query += `\nAND campaign.name LIKE "%${CAMPAIGN_FILTER}%"`;
    }
    
    // Add campaign exclusion filter if specified
    if (CAMPAIGN_EXCLUSION_FILTER && CAMPAIGN_EXCLUSION_FILTER.trim() !== '') {
        query += `\nAND campaign.name NOT LIKE "%${CAMPAIGN_EXCLUSION_FILTER}%"`;
    }
    
    query += `\nORDER BY metrics.cost_micros DESC`;
    
    return query;
}

function processSearchTermsWithCampaignCheck(searchTermRows) {
    const data = [];
    let processedCount = 0;
    let qualifyingCount = 0;
    
    // Cache to store campaign keywords to avoid repeated queries
    const campaignKeywordsCache = new Map();
    
    while (searchTermRows.hasNext()) {
        try {
            const row = searchTermRows.next();
            processedCount++;
            
            // Access fields using dot notation
            const searchTerm = row.searchTermView && row.searchTermView.searchTerm ? row.searchTermView.searchTerm : '';
            const status = row.searchTermView && row.searchTermView.status ? row.searchTermView.status : '';
            const campaignId = row.campaign && row.campaign.id ? row.campaign.id : '';
            const campaignName = row.campaign && row.campaign.name ? row.campaign.name : '';
            const adGroupId = row.adGroup && row.adGroup.id ? row.adGroup.id : '';
            const adGroupName = row.adGroup && row.adGroup.name ? row.adGroup.name : '';
            
            // Convert metrics to numbers
            const impressions = Number(row.metrics && row.metrics.impressions ? row.metrics.impressions : 0);
            const clicks = Number(row.metrics && row.metrics.clicks ? row.metrics.clicks : 0);
            const costMicros = Number(row.metrics && row.metrics.costMicros ? row.metrics.costMicros : 0);
            const conversions = Number(row.metrics && row.metrics.conversions ? row.metrics.conversions : 0);
            const conversionValue = Number(row.metrics && row.metrics.conversionsValue ? row.metrics.conversionsValue : 0);
            
            // Calculate derived metrics
            const cost = costMicros / 1000000; // Convert micros to currency
            const cpa = conversions > 0 ? cost / conversions : 0;
            const roas = cost > 0 ? conversionValue / cost : 0;
            
            // Check if search term is already a keyword in any ad group of this campaign
            const isAlreadyKeyword = checkIfSearchTermIsKeywordInCampaign(searchTerm, campaignId, campaignKeywordsCache);
            
            // Check if search term meets criteria
            if (!isAlreadyKeyword && meetsThresholds(cost, clicks, conversions, cpa, roas, status)) {
                qualifyingCount++;
                
                const newRow = [
                    searchTerm,
                    status,
                    campaignName,
                    adGroupName,
                    impressions,
                    clicks,
                    cost,
                    conversions,
                    conversionValue,
                    cpa,
                    roas,
                    '', // Placeholder for AI Classification
                    ''  // Placeholder for AI Reasoning
                ];
                
                data.push(newRow);
            }
            
            // Log progress every 1000 rows
            if (processedCount % 1000 === 0) {
                Logger.log(`Processed ${processedCount} rows, found ${qualifyingCount} qualifying terms so far`);
            }
            
        } catch (e) {
            Logger.log(`Error processing row ${processedCount}: ${e}`);
            // Continue with next row
        }
    }
    
    Logger.log(`Processed ${processedCount} search terms, found ${qualifyingCount} qualifying terms`);
    return data;
}

function checkIfSearchTermIsKeywordInCampaign(searchTerm, campaignId, cache) {
    // Check cache first
    if (cache.has(campaignId)) {
        const campaignKeywords = cache.get(campaignId);
        return campaignKeywords.includes(searchTerm.toLowerCase());
    }
    
    // If not in cache, get all keywords for this campaign
    const campaignKeywords = getAllKeywordsInCampaign(campaignId);
    cache.set(campaignId, campaignKeywords);
    
    return campaignKeywords.includes(searchTerm.toLowerCase());
}

function getAllKeywordsInCampaign(campaignId) {
    const keywords = [];
    
    try {
        // Get campaign name first
        const campaign = AdsApp.campaigns()
            .withCondition(`campaign.id = ${campaignId}`)
            .get()
            .next();
        const campaignName = campaign.getName();
        
        // Check if campaign should be excluded
        if (CAMPAIGN_EXCLUSION_FILTER && CAMPAIGN_EXCLUSION_FILTER.trim() !== '' && 
            campaignName.toLowerCase().includes(CAMPAIGN_EXCLUSION_FILTER.toLowerCase())) {
            Logger.log(`Skipping excluded campaign "${campaignName}" (contains "${CAMPAIGN_EXCLUSION_FILTER}")`);
            return keywords;
        }
        
        // Get all ad groups in this campaign
        const adGroups = AdsApp.adGroups()
            .withCondition(`campaign.id = ${campaignId}`)
            .get();
        
        while (adGroups.hasNext()) {
            const adGroup = adGroups.next();
            
            // Get all keywords in this ad group
            const adGroupKeywords = adGroup.keywords()
                .withCondition('Status = ENABLED')
                .get();
            
            while (adGroupKeywords.hasNext()) {
                const keyword = adGroupKeywords.next();
                keywords.push(keyword.getText().toLowerCase());
            }
        }
        
        Logger.log(`Found ${keywords.length} keywords in campaign "${campaignName}"`);
        
    } catch (e) {
        Logger.log(`Error getting keywords for campaign ${campaignId}: ${e}`);
    }
    
    return keywords;
}

function meetsThresholds(cost, clicks, conversions, cpa, roas, status) {
    // Check if already added as keyword (status = 'ADDED') or excluded as negative (status = 'EXCLUDED')
    if (status === 'ADDED' || status === 'EXCLUDED') {
        return false;
    }
    
    // Check basic thresholds
    if (cost < MIN_COST) return false;
    if (clicks < MIN_CLICKS) return false;
    if (conversions < MIN_CONVERSIONS) return false;
    
    // Check target threshold
    if (TARGET_TYPE === 'CPA') {
        if (cpa > TARGET_VALUE) return false; // CPA should be lower than target
    } else if (TARGET_TYPE === 'ROAS') {
        if (roas < TARGET_VALUE) return false; // ROAS should be higher than target
    }
    
    return true;
}

function applyOpenAIClassification(data) {
    if (!USE_OPENAI_CLASSIFICATION || data.length === 0) {
        return data;
    }
    
    try {
        // Get OpenAI API key
        const apiKey = getOpenAIAPIKey();
        if (!apiKey) {
            Logger.log('OpenAI API key not found. Skipping AI classification.');
            return data;
        }
        
        // Get website content if URL is provided
        const websiteContent = WEBSITE_URL ? getWebsiteContent(WEBSITE_URL) : '';
        
        // Estimate costs before processing
        const estimatedCost = estimateOpenAICosts(data.length, websiteContent);
        Logger.log(`Estimated OpenAI cost: $${estimatedCost.toFixed(4)}`);
        
        if (estimatedCost > MAX_TOTAL_COST) {
            Logger.log(`Estimated cost ($${estimatedCost.toFixed(4)}) exceeds maximum ($${MAX_TOTAL_COST}). Skipping AI classification.`);
            return data;
        }
        
        // Process search terms in batches
        for (let i = 0; i < data.length; i += BATCH_SIZE) {
            const batch = data.slice(i, i + BATCH_SIZE);
            const batchResults = classifySearchTermsBatch(batch, apiKey, websiteContent);
            
            // Update the data with classification results
            for (let j = 0; j < batchResults.length; j++) {
                if (i + j < data.length) {
                    data[i + j][11] = batchResults[j].classification; // AI Classification
                    data[i + j][12] = batchResults[j].reasoning;      // AI Reasoning
                }
            }
            
            // Add delay between batches to avoid rate limits
            if (i + BATCH_SIZE < data.length) {
                Utilities.sleep(1000); // 1 second delay
            }
            
            Logger.log(`Processed batch ${Math.floor(i/BATCH_SIZE) + 1}/${Math.ceil(data.length/BATCH_SIZE)}`);
        }
        
        Logger.log(`AI classification completed for ${data.length} search terms`);
        
    } catch (e) {
        Logger.log(`Error in AI classification: ${e.message}`);
        Logger.log('Continuing without AI classification...');
    }
    
    return data;
}

function getOpenAIAPIKey() {
    try {
        const ss = SpreadsheetApp.openByUrl(SHEET_URL);
        const apiKeyRange = ss.getRangeByName(OPENAI_API_KEY_NAMED_RANGE);
        if (!apiKeyRange) {
            Logger.log(`Named range '${OPENAI_API_KEY_NAMED_RANGE}' not found in the spreadsheet`);
            return null;
        }
        return apiKeyRange.getValue();
    } catch (e) {
        Logger.log(`Error getting OpenAI API key: ${e.message}`);
        return null;
    }
}

function getWebsiteContent(url) {
    try {
        const response = UrlFetchApp.fetch(url, {
            muteHttpExceptions: true,
            followRedirects: true
        });
        
        if (response.getResponseCode() === 200) {
            const html = response.getContentText();
            
            // Extract text content (simplified - could be enhanced)
            const textContent = html
                .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '') // Remove scripts
                .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')   // Remove styles
                .replace(/<[^>]+>/g, ' ')                          // Remove HTML tags
                .replace(/\s+/g, ' ')                              // Normalize whitespace
                .trim()
                .substring(0, 2000); // Limit to 2000 characters
            
            Logger.log(`Extracted ${textContent.length} characters from website`);
            return textContent;
        } else {
            Logger.log(`Failed to fetch website content. Response code: ${response.getResponseCode()}`);
            return '';
        }
    } catch (e) {
        Logger.log(`Error fetching website content: ${e.message}`);
        return '';
    }
}

function estimateOpenAICosts(searchTermCount, websiteContent) {
    const avgTokensPerTerm = 50; // Average tokens per search term
    const websiteTokens = websiteContent.length / 4; // Rough estimate: 4 characters per token
    const totalInputTokens = (avgTokensPerTerm * searchTermCount) + websiteTokens;
    const totalOutputTokens = searchTermCount * 20; // Average 20 tokens per response
    
    const pricing = OPENAI_PRICING[OPENAI_MODEL];
    const inputCost = (totalInputTokens / 1000) * pricing.input;
    const outputCost = (totalOutputTokens / 1000) * pricing.output;
    
    return inputCost + outputCost;
}

function classifySearchTermsBatch(searchTerms, apiKey, websiteContent) {
    const results = [];
    
    try {
        // Create prompt for the batch
        const prompt = createClassificationPrompt(searchTerms, websiteContent);
        
        // Call OpenAI API
        const response = callOpenAIAPI(prompt, apiKey);
        
        // Parse response
        const classifications = parseClassificationResponse(response, searchTerms.length);
        
        return classifications;
        
    } catch (e) {
        Logger.log(`Error classifying batch: ${e.message}`);
        // Return default classifications for this batch
        return searchTerms.map(() => ({
            classification: 'REVIEW',
            reasoning: 'Error in AI classification - manual review required'
        }));
    }
}

function createClassificationPrompt(searchTerms, websiteContent) {
    const searchTermsList = searchTerms.map(term => term[0]).join('\n- ');
    
    let prompt = `Analyze the relevance of these search terms to this business:

SEARCH TERMS:
- ${searchTermsList}

`;

    if (websiteContent) {
        prompt += `WEBSITE CONTENT:
${websiteContent}...

`;
    }

    prompt += `For each search term, classify as:
- RELEVANT: Directly related to products/services
- SEMI_RELEVANT: Somewhat related, could be valuable
- IRRELEVANT: Not related to business
- COMPETITOR: Competitor brand names
- GENERIC: Too broad/generic terms

Respond in this exact format for each term:
TERM: [search term]
CLASSIFICATION: [classification]
REASONING: [brief explanation]

`;

    return prompt;
}

function callOpenAIAPI(prompt, apiKey) {
    const payload = {
        model: OPENAI_MODEL,
        messages: [
            {
                role: "system",
                content: "You are a search term classifier for Google Ads. Analyze search terms for business relevance."
            },
            {
                role: "user",
                content: prompt
            }
        ],
        temperature: 0.3,
        max_tokens: 1000
    };
    
    const options = {
        method: "post",
        headers: {
            "Authorization": `Bearer ${apiKey}`,
            "Content-Type": "application/json"
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.error) {
        throw new Error(`OpenAI API error: ${responseData.error.message}`);
    }
    
    // Track usage and calculate cost
    const usage = responseData.usage;
    totalInputTokens += usage.prompt_tokens;
    totalOutputTokens += usage.completion_tokens;
    apiCallCount++;
    
    const pricing = OPENAI_PRICING[OPENAI_MODEL];
    const inputCost = (usage.prompt_tokens / 1000) * pricing.input;
    const outputCost = (usage.completion_tokens / 1000) * pricing.output;
    const callCost = inputCost + outputCost;
    totalCost += callCost;
    
    Logger.log(`API call ${apiCallCount}: $${callCost.toFixed(4)} (${usage.prompt_tokens} input, ${usage.completion_tokens} output tokens)`);
    
    return responseData.choices[0].message.content;
}

function parseClassificationResponse(response, expectedCount) {
    const results = [];
    const lines = response.split('\n');
    
    let currentTerm = '';
    let currentClassification = '';
    let currentReasoning = '';
    
    for (const line of lines) {
        if (line.startsWith('TERM:')) {
            // Save previous result if exists
            if (currentTerm && currentClassification) {
                results.push({
                    classification: currentClassification,
                    reasoning: currentReasoning
                });
            }
            
            currentTerm = line.replace('TERM:', '').trim();
            currentClassification = '';
            currentReasoning = '';
        } else if (line.startsWith('CLASSIFICATION:')) {
            currentClassification = line.replace('CLASSIFICATION:', '').trim();
        } else if (line.startsWith('REASONING:')) {
            currentReasoning = line.replace('REASONING:', '').trim();
        }
    }
    
    // Add the last result
    if (currentTerm && currentClassification) {
        results.push({
            classification: currentClassification,
            reasoning: currentReasoning
        });
    }
    
    // Ensure we have the right number of results
    while (results.length < expectedCount) {
        results.push({
            classification: 'REVIEW',
            reasoning: 'Failed to parse AI response - manual review required'
        });
    }
    
    return results.slice(0, expectedCount);
}

function logCostSummary() {
    if (USE_OPENAI_CLASSIFICATION && apiCallCount > 0) {
        Logger.log('=== OPENAI COST SUMMARY ===');
        Logger.log(`Total API calls: ${apiCallCount}`);
        Logger.log(`Total input tokens: ${totalInputTokens}`);
        Logger.log(`Total output tokens: ${totalOutputTokens}`);
        Logger.log(`Total cost: $${totalCost.toFixed(4)}`);
        Logger.log(`Average cost per search term: $${(totalCost / (apiCallCount * BATCH_SIZE)).toFixed(6)}`);
        Logger.log('===========================');
    }
}

function exportToSheet(data) {
    try {
        // Handle spreadsheet
        let ss;
        if (!SHEET_URL) {
            Logger.log("Creating new spreadsheet...");
            const accountName = AdsApp.currentAccount().getName();
            const spreadsheetName = `Keyword Expansion Script - ${accountName}`;
            ss = SpreadsheetApp.create(spreadsheetName);
            const url = ss.getUrl();
            Logger.log("Created new spreadsheet: " + url);
        } else {
            Logger.log("Opening existing spreadsheet...");
            ss = SpreadsheetApp.openByUrl(SHEET_URL);
        }
        
        // Create or clear the sheet
        let sheet;
        if (ss.getSheetByName(TAB)) {
            sheet = ss.getSheetByName(TAB);
            sheet.clear();
            Logger.log(`Cleared existing sheet: ${TAB}`);
        } else {
            sheet = ss.insertSheet(TAB);
            Logger.log(`Created new sheet: ${TAB}`);
        }
        
        // Create headers
        const headers = [
            'Search Term',
            'Status',
            'Campaign',
            'Ad Group',
            'Impressions',
            'Clicks',
            'Cost',
            'Conversions',
            'Conv. Value',
            'CPA',
            'ROAS',
            'AI Classification',
            'AI Reasoning'
        ];
        
        // Write headers and data to sheet in a single operation
        if (data.length > 0) {
            Logger.log(`Writing ${data.length} rows to spreadsheet...`);
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(2, 1, data.length, headers.length).setValues(data);
            Logger.log(`Successfully wrote ${data.length} rows of data to the spreadsheet.`);
        } else {
            Logger.log("No qualifying search terms found.");
            sheet.getRange(1, 1).setValue("No qualifying search terms found for the specified criteria.");
        }
        
    } catch (e) {
        Logger.log(`Error in exportToSheet: ${e.message}`);
        Logger.log("Attempting to log data to console instead...");
        
        // Fallback: log the data to console
        if (data.length > 0) {
            Logger.log("=== SEARCH TERM DATA ===");
            Logger.log("Search Term | Status | Campaign | Ad Group | Impressions | Clicks | Cost | Conversions | Conv. Value | CPA | ROAS");
            data.forEach(row => {
                Logger.log(row.join(" | "));
            });
            Logger.log("=== END DATA ===");
        } else {
            Logger.log("No qualifying search terms found for the specified criteria.");
        }
        
        throw e; // Re-throw the error so the main function knows something went wrong
    }
}