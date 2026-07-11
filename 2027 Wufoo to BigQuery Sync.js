/**
 * Main function to sync 5 Wufoo forms to 5 BigQuery tables via Upsert (Update/Insert).
 */
function syncWufooToBigQuery() {
  // 1. Core Configuration
  const WUFOO_API_KEY = 'VHA4-QIF5-FUJ1-97EE';
  const WUFOO_SUBDOMAIN = 'albsok'; 
  const BQ_PROJECT_ID = 'session-master-data';
  const BQ_DATASET_ID = 'wufoo_data';
  
  // 2. Map Wufoo Forms to BigQuery Tables and define custom schemas
  // System fields (EntryId, DateCreated, Status) are intentionally omitted here, as they are hardcoded below.
  const configurations = [
    { 
      formId: 'kansas-boys-state-2027-online-application', 
      tableId: 'app_data',
      fieldMap: {
        "Welcome Letter Sent": "welcome_letter_sent",
        "Status_669": "status", // Custom radio field
        "Carryover?": "carryover",
        "Early Bird": "early_bird",
        "Total Paid": "total_paid",
        "Sponsor Info": "sponsor_info",
        "First_1516": "delegate_first_name",
        "Last_1517": "delegate_last_name",
        "Last_1514": "parent_last_name",
        "Last_1512": "emergency_contact_last_name",
        "Delegate Mobile Phone Number": "delegate_mobile_phone_number", 
        "Delegate Email Address": "delegate_email_address",
        "Delegate Date of Birth": "delegate_dob",
        "Delegate High School (A-L)": "high_school_al",
        "Delegate High School (M-Z)": "high_school_mz",
        "Other School": "other_school",
        "Delegate High School (Not Listed)": "high_school_not_listed",
        "Delegate Grade Level (during 2026-2027 school year)": "grade_level",
        "Will this be your first or second time attending Kansas Boys State?": "first_or_second_time_attending",
        "Delegate Address": "delegate_address",
        "Delegate City": "delegate_city",
        "Delegate State": "delegate_state",
        "Delegate Zip Code": "delegate_zip_code",
        "Delegate T-Shirt Size": "tshirt_size",
        "It strength/weight training required by your coach?": "conditioning_required",
        "Are you interested in Boy Scout Merit Badge opportunities?": "merit_badge_interest",
        "What is your rank?": "scout_rank",
        "American Heritage_1291": "american_heritage_interest",
        "Citizenship in Community_1292": "community_interest",
        "Citizenship in Nation_1293": "nation_interest",
        "Citizenship in Society_1294": "society_interest",
        "Communications_1295": "communications_interest",
        "Law_1296": "law_interest",
        "Public Speaking_1297": "public_speaking_interest",
        "Scout Master Name": "scout_master_name",
        "Troop Number": "troop_number",
        "First_1513": "parent_first_name",
        "Parent/Guardian Email": "parent_email",
        "Parent/Guardian Phone Number": "parent_phone_number",
        "First_1511": "emergency_contact_first_name",
        "Emergency Contact Phone Number": "emergency_contact_phone_number",
        "Emergency Contact Relationship": "emergency_contact_relationship",
        "Special Needs/Accomocations Needed": "accommodations_needed",
        "Anything else we should know about you?": "any_other_info",
        "Do you want to pay the $50 delegate fee now?": "pay_fee_now",
        "Payment Status": "payment_status",
        "Payment Total": "payment_total",
        "Payment Currency": "payment_currency",
        "Payment Confirmation": "payment_confirmation",
        "Payment Merchant": "payment_merchant"
      }
    },
    { 
      formId: '2027-delegate-social-contract', 
      tableId: 'social_contract',
      fieldMap: {
        "Delegate Application ID#": "application_id",
        "First_2": "delegate_first_name",
        "Last_3": "delegate_last_name",
        "Email": "delegate_email"
      }
    },
    { 
      formId: '2027-stater-medical-information-and-media-consent', 
      tableId: 'medical_media_consent',
      fieldMap: {
        "Delegate Application ID#": "application_id",
        "First_123": "delegate_first_name",
        "Last_124": "delegate_last_name",
        "Parent/Guardian Email": "parent_email",
        "Entering your name as parent/guardian serves as consent for emergency medical treatment for the above named delegate": "medical_consent_name",
        "Entering your name as parent/guardian serves as consent for use of all media and demographic information for the above named delegate, captured during the  American Legion Boys State of Kansas session, to be used in legitimate Boys State promotions without rights of compensation or ownership.": "media_consent_name",
        "Food, medication, or other allergies": "allergies_select",
        "Current prescription medications": "prescriptions_select",
        "Assistance or storage of medication needs": "medication_needs_select",
        "Conditions limiting physical activities": "physical_needs_select",
        "Current Medical Insurance": "medical_insurance_select",
        "List food, medication, or other allergies": "allergy_info",
        "Are your allergies anaphylactic?": "anaphylactic_allergies",
        "Soy/Soybeans": "soy",
        "Eggs": "eggs",
        "Milk/Lactose": "lactose",
        "Sesame": "sesame",
        "Tree Nuts": "treenuts",
        "Wheat/Gluten": "gluten",
        "Fish": "fish",
        "Peanuts": "peanuts",
        "Shellfish": "shellfish",
        "List current prescription medications and dosage": "medications",
        "Describe assistance or storage of medication needs": "med_storage",
        "List conditions limiting physical activities (e.g. athletics, walking/standing for long periods)": "physical_conditions",
        "Medical Insurance Carrier": "medical_insurance_carrier",
        "Policy Number": "medical_insurance_policy"
      }
    },
    { 
      formId: '2027-boys-state-precheckin-meal', 
      tableId: 'precheckin_meal',
      fieldMap: {
        "Delegate Application ID#": "application_id",
        "First": "delegate_first_name",
        "Last": "delegate_last_name",
        "Email to send confirmation": "email_confirmation",
        "Diners over 10 years old": "count_over_10",
        "Diners 5-10 years old": "count_5_to_10"
      }
    },
    { 
      formId: '2027-boys-state-parking-info', 
      tableId: 'delegate_parking',
      fieldMap: {
        "Delegate Application ID#": "application_id",
        "First_3": "delegate_first_name",
        "Last_4": "delegate_last_name",
        "Email": "delegate_email", 
        "Delegate Email": "delegate_email", 
        "Country of License Plate": "country",
        "State of License Plate": "state",
        "Number on License Plate": "number",
        "Notes": "notes"
      }
    },
  ];

  // 3. Loop through and process each configuration
  configurations.forEach(config => {
    try {
      // Fetch data from Wufoo
      const wufooData = fetchWufooData(WUFOO_SUBDOMAIN, WUFOO_API_KEY, config.formId);
      
      // If entries exist, process and push to BigQuery
      if (wufooData && wufooData.Entries && wufooData.Entries.length > 0) {
        
        // Dynamic background translation map step to compile FieldXX -> Friendly Label strings
        const translationMap = buildWufooTranslationMap(WUFOO_SUBDOMAIN, WUFOO_API_KEY, config.formId, config.fieldMap);
        
        // Translate raw Wufoo keys to your specified BigQuery column names
        const cleanedEntries = wufooData.Entries.map(entry => {
          const normalizedEntry = {};
          
          // Map raw API fields into friendly code labels before translating schemas
          for (const rawKey in entry) {
            const codeSetKey = translationMap[rawKey] || rawKey;
            normalizedEntry[codeSetKey] = entry[rawKey];
          }
          
          // Execute custom fields mapping
          const bqEntry = mapWufooToBigQuerySchema(normalizedEntry, config.fieldMap, config.tableId);
          
          // CRITICAL DIRECT INJECTION: Bypass translation engine and force system variables locally
          // Covers both upper and lower-case variants produced by different versions of the Wufoo REST payload
          bqEntry['entry_id'] = entry['EntryId'] || entry['entry_id'] || entry['id'];
          bqEntry['date_created'] = entry['DateCreated'] || entry['date_created'];
          
          if (config.tableId === 'app_data') {
            bqEntry['payment_status'] = entry['Status'] || entry['status'] || entry['PaymentStatus'] || ""; 
            bqEntry['payment_total'] = entry['PurchaseTotal'] || entry['purchasetotal'] || entry['PaymentTotal'] || 0;
            bqEntry['application_id'] = String(bqEntry['entry_id']).padStart(6, '0');
          }
          
          if (config.tableId === 'precheckin_meal') {
            bqEntry['payment_status'] = entry['Status'] || entry['status'] || entry['PaymentStatus'] || "";
            bqEntry['payment_total'] = entry['PurchaseTotal'] || entry['purchasetotal'] || entry['PaymentTotal'] || 0;
          }
          
          // Post-Injection: Calculate boolean completion tracking explicitly using the guaranteed entry_id
          const booleanCompletedTables = [
            'delegate_parking', 'medical_media_consent', 'precheckin_meal', 'social_contract'
          ];
          if (booleanCompletedTables.includes(config.tableId)) {
            const entryIdNum = Number(bqEntry['entry_id']);
            bqEntry['completed'] = (!isNaN(entryIdNum) && entryIdNum > 0);
          }
          
          return bqEntry;
        });
        
        // UPSERT into BigQuery (Creates new, updates existing)
        upsertToBigQuery(BQ_PROJECT_ID, BQ_DATASET_ID, config.tableId, cleanedEntries);
        Logger.log(`Success: Upserted ${wufooData.Entries.length} rows from ${config.formId} to ${config.tableId}`);
      } else {
        Logger.log(`Notice: No data found for form ${config.formId}`);
      }
    } catch (e) {
      Logger.log(`Error syncing ${config.formId}: ${e.message}`);
    }
  });
}

/**
 * Helper: Translates keys based on the specific fieldMap provided for the form.
 */
function mapWufooToBigQuerySchema(entry, fieldMap, tableId) {
  const mappedEntry = {};
  
  const yesNoStringColumns = [
    "american_heritage_interest", "community_interest", "nation_interest", 
    "society_interest", "communications_interest", "law_interest", 
    "public_speaking_interest", "soy", "eggs", "lactose", "sesame", 
    "treenuts", "gluten", "fish", "peanuts", "shellfish",
    "early_bird" 
  ];

  const booleanColumns = [
    "welcome_letter_sent", "conditioning_required", "anaphylactic_allergies",
    "carryover", "pay_fee_now", "merit_badge_interest"
  ];
  
  for (const wufooKey in entry) {
    if (entry.hasOwnProperty(wufooKey)) {
      
      const bqColumnName = fieldMap[wufooKey] || wufooKey;
      let value = entry[wufooKey];
      
      // --- PADDING LOGIC (for application_id fields outside of app_data) ---
      if (bqColumnName === "application_id") {
        if (value && value.toString().trim() !== "") {
          value = String(value).padStart(6, '0');
        }
      }
      
      // Handle Strict Booleans
      if (booleanColumns.includes(bqColumnName)) {
        const strVal = value ? value.toString().toLowerCase().trim() : "";
        value = (strVal === "yes" || strVal === "true" || strVal === "1");
      }
      // Handle Yes/No Strings
      else if (yesNoStringColumns.includes(bqColumnName)) {
        if (value && value.toString().trim() !== "") {
          value = "Yes";
        } else {
          value = "No";
        }
      }
      
      mappedEntry[bqColumnName] = value;
    }
  }
  
  // Safety Catch: Ensure missing Yes/No string fields get "No"
  yesNoStringColumns.forEach(col => {
    if (!mappedEntry.hasOwnProperty(col)) {
      mappedEntry[col] = "No";
    }
  });

  // Safety Catch: Ensure missing Strict Boolean fields get false
  booleanColumns.forEach(col => {
    if (!mappedEntry.hasOwnProperty(col)) {
      mappedEntry[col] = false;
    }
  });

  // --- SAFETY NET FOR REQUIRED STRINGS ---
  if (tableId === 'medical_media_consent' && !mappedEntry.hasOwnProperty('delegate_last_name')) {
    mappedEntry['delegate_last_name'] = "";
  }
  if (tableId === 'delegate_parking' && !mappedEntry.hasOwnProperty('delegate_email')) {
    mappedEntry['delegate_email'] = "";
  }
  
  return mappedEntry;
}

/**
 * Background Engine: Maps Wufoo's API alphanumeric field identifiers (Field1, Field43)
 */
function buildWufooTranslationMap(subdomain, apiKey, formId, fieldMap) {
  const url = `https://${subdomain}.wufoo.com/api/v3/forms/${formId}/fields.json?system=true`;
  const options = {
    method: 'get',
    headers: { "Authorization": "Basic " + Utilities.base64Encode(apiKey + ':footastic') },
    muteHttpExceptions: true
  };
  
  const translationMap = {};
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data && data.Fields) {
        const codeSetKeys = Object.keys(fieldMap);
        
        data.Fields.forEach(field => {
          const bqKey = findCodesetKey(field.ID, field.Title, null, codeSetKeys);
          if (bqKey) translationMap[field.ID] = bqKey;
          
          if (field.SubFields && field.SubFields.length > 0) {
            field.SubFields.forEach(sub => {
              const subBqKey = findCodesetKey(sub.ID, field.Title, sub.Label, codeSetKeys);
              if (subBqKey) translationMap[sub.ID] = subBqKey;
            });
          }
        });
      }
    }
  } catch (e) {
    Logger.log(`Warning: Field mapping initialization skipped for ${formId}: ${e.message}`);
  }
  return translationMap;
}

/**
 * Priority Lookup Helper
 */
function findCodesetKey(fieldId, title, label, codeSetKeys) {
  const idNum = fieldId.replace('Field', '');
  const cleanTitle = title ? title.trim() : "";
  const cleanLabel = label ? label.trim() : "";
  
  if (cleanLabel) {
    const match = codeSetKeys.find(k => k.trim() === `${cleanLabel}_${idNum}`);
    if (match) return match;
  }
  if (cleanTitle) {
    const match = codeSetKeys.find(k => k.trim() === `${cleanTitle}_${idNum}`);
    if (match) return match;
  }
  if (cleanTitle) {
    const match = codeSetKeys.find(k => k.trim() === cleanTitle);
    if (match) return match;
  }
  if (cleanLabel) {
    const match = codeSetKeys.find(k => k.trim() === cleanLabel);
    if (match) return match;
  }
  return null;
}

/**
 * Evaluates form entries using the Wufoo API.
 */
function fetchWufooData(subdomain, apiKey, formId) {
  const url = `https://${subdomain}.wufoo.com/api/v3/forms/${formId}/entries.json?system=true`;
  const options = {
    method: 'get',
    headers: { "Authorization": "Basic " + Utilities.base64Encode(apiKey + ':footastic') },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    return JSON.parse(response.getContentText());
  } else {
    throw new Error(`Wufoo API Error (${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

/**
 * Upserts data into BigQuery by deleting existing matched entry_ids first, 
 * then batch-loading the new payloads. This ensures rows are updated without duplicates.
 */
function upsertToBigQuery(projectId, datasetId, tableId, entries) {
  if (!entries || entries.length === 0) return;

  // 1. Delete Existing Records (The "Update" half of the Upsert)
  const entryIds = entries.map(e => e.entry_id).filter(id => id != null);
  
  if (entryIds.length > 0) {
    const formattedIds = entryIds.map(id => `'${id}'`).join(',');
    const deleteQuery = `DELETE FROM \`${projectId}.${datasetId}.${tableId}\` WHERE CAST(entry_id AS STRING) IN (${formattedIds})`;
    
    const queryRequest = {
      query: deleteQuery,
      useLegacySql: false
    };
    
    try {
      BigQuery.Jobs.query(queryRequest, projectId);
      Logger.log(`Prepared ${tableId} for updates (cleared matched old records).`);
    } catch (e) {
      Logger.log(`Notice: DELETE step skipped for ${tableId}. (Table may be empty). Msg: ${e.message}`);
    }
  }

  // 2. Load the new/updated records using a Batch Load (The "Insert" half of the Upsert)
  const ndjson = entries.map(e => JSON.stringify(e)).join('\n');
  const blob = Utilities.newBlob(ndjson, 'application/octet-stream');
  
  const loadJob = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        sourceFormat: 'NEWLINE_DELIMITED_JSON',
        writeDisposition: 'WRITE_APPEND',
        ignoreUnknownValues: true 
      }
    }
  };
  
  try {
    BigQuery.Jobs.insert(loadJob, projectId, blob);
  } catch (e) {
    throw new Error(`BigQuery Upsert Load Failure for ${tableId}: ${e.message}`);
  }
}