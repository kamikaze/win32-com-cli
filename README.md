# win32-com-cli
CLI tool to communicate with Win32 COM


Sample structure to pass via stdin:
```json
{
  "version": "1",
  "prog_id": "ECR2ATL.ECR2Transaction",
  "method": "Cancellation",
  "properties": {
    "ECRNameAndVersion": "App Ver. 123.321",
    "ReqInvoiceNumber": "NR12345",
    "ReqDateTime": "2025-05-22 12:33:44"
  }
}
```