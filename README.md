# EngageScoutLambda

AWS Lambda function for automating scout messages on the Engage recruitment platform.

## Overview

This Lambda function automates the process of:
- Logging into Engage platform
- Filtering candidates by age and prefecture
- Sending scout messages automatically

## Quick Start

1. Install dependencies:
   ```bash
   npm install
   ```

2. Test locally:
   ```bash
   node test-lambda-locally.js
   ```

3. Package for deployment:
   ```bash
   ./package_function.sh
   ```

## Request Format

```json
{
  "id": "your-email@example.com",
  "pass": "your-password",
  "min_age": 25,
  "max_age": 35,
  "prefectures": ["東京都", "神奈川県"]
}
```

## Documentation

- [Lambda Setup Guide](README_LAMBDA.md)
- [Sample Requests](sample-requests/README.md)
- [Development Guide](CLAUDE.md)

## Architecture

- **Runtime**: Node.js 18.x with Puppeteer
- **Deployment**: AWS Lambda with Chrome layers
- **Input**: JSON via POST request
- **Output**: Scout results with count

## License

ISC