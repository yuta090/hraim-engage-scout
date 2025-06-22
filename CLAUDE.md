# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

EngageScoutLambda is an AWS Lambda function that automates scout message sending on the Engage recruitment platform using Puppeteer. The project has transitioned from a Flet GUI desktop application to a serverless Lambda function that receives parameters via POST requests.

## Key Architecture Decisions

1. **Puppeteer over Selenium**: The project uses Puppeteer instead of Selenium for better Lambda compatibility and smaller deployment size.

2. **Environment Detection**: The code automatically detects Lambda vs local environment:
   ```javascript
   const isLambdaEnvironment = !!(process.env.AWS_LAMBDA_FUNCTION_NAME || process.env.LAMBDA_RUNTIME_DIR);
   ```

3. **Lambda Layer Dependencies**: Uses @sparticuz/chromium for Lambda-optimized Chromium binary, avoiding the need to bundle Chrome in the deployment package.

## Common Development Commands

```bash
# Install dependencies
npm install

# Run local test with sample data
node test-lambda-locally.js

# Test specific request pattern (1-4)
node test-lambda-locally.js 1

# Package for Lambda deployment
./package_function.sh
# Creates: dist/EngageScoutLambda.zip

# Direct local test with custom data
node local-test.js
```

## Core Modules Structure

- **index.js**: Lambda handler that manages browser lifecycle and request validation
- **lib/engage-scout.js**: Core automation logic for Engage platform
  - Login flow with PK token extraction
  - Candidate filtering by age and prefecture
  - Scout message sending with rate limiting

## Request/Response Format

```json
// Request
{
  "id": "login-email@example.com",    // Required
  "pass": "password",                  // Required  
  "min_age": 25,                      // Optional (default: 21)
  "max_age": 35,                      // Optional (default: 60)
  "prefectures": ["東京都", "神奈川県"] // Optional (default: all)
}

// Response
{
  "scoutCount": 10,
  "processedCount": 15,
  "status": "completed|error",
  "error": "error message if failed"
}
```

## Lambda Deployment Requirements

1. **Required Lambda Layers**:
   - puppeteer-core-layer
   - sparticuz-chromium-layer  
   - chromium-fonts-layer (for Japanese text)

2. **Lambda Configuration**:
   - Runtime: Node.js 18.x
   - Memory: 1024 MB minimum
   - Timeout: 5 minutes minimum
   - Architecture: x86_64 (required for Chromium)

## Key Implementation Details

1. **Browser Configuration**: Different settings for Lambda vs local environments in `getBrowserConfig()`

2. **Error Handling**: Comprehensive try-catch with browser cleanup in finally block

3. **Processing Limits**: Hard limit of 50 candidates per execution to avoid Lambda timeout

4. **Rate Limiting**: Random wait between 4-5 seconds between actions using `randomWait()`

5. **Selector Patterns**: Multiple fallback selectors for robustness against UI changes

## Testing Approach

- Use `sample-requests/` directory for different test scenarios
- Local testing simulates Lambda event structure
- Test both minimal (id/pass only) and full parameter sets
- Always test login flow first before full scout process

## Legacy Code Notes

The `docs/scout_related_files/` directory contains the original Flet GUI implementation using Selenium. This code is preserved for reference but is not used in the Lambda function.