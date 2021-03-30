Imports Microsoft.VisualBasic

Public Enum pakPOStates
  Free = 1
  UnderSupplierVerification = 2
  UnderISGECApproval = 3
  IssuedtoSupplier = 4
  UnderDespatch = 5
  Closed = 6
  ImportedFromERP = 7
End Enum
Public Enum pakAlertEvents
  POIssued = 1
  DocumentsDespatched = 2
  MaterialDespatched = 3
  POClosed = 4
  OpenPORequested = 5
  OpenPORequestExecuted = 6
  MaterialReceiptAtSite = 7
  POVerification = 8
  POApproval = 9
  TCPOIssued = 10
End Enum
Public Enum pakItemStates
  CreatedByISGEC = 1
  CreatedBySupplier = 2
  VerifiedByISGEC = 3
  VerifiedbySupplier = 4
  ChangeRequiredByISGEC = 5
  ChangeRequiredBySupplier = 6
  AcceptedbyISGEC = 7
  AcceptedbySupplier = 8
  FreezedbyISGEC = 9
  FreezedbySupplier = 10
End Enum
Public Enum pakTCPOStates
  Free = 1
  IssuedToSupplier = 4
  Closed = 6
  ReopenRequested = 7
End Enum
Public Enum pakTCPOLineStates
  Free = 1
  TCCreated = 2
  TCSubmitted = 3
  TCCommented = 4
  TCCleared = 5
End Enum
Public Enum pakTCUploadStates
  Free = 1
  UnderEvaluation = 2
  CommentSubmitted = 3
  TechnicallyCleared = 4
  Closed = 5
  Superseded = 6
End Enum
Public Enum pakTCAlertEvents
  DocumentsSubmitted = 2
  POClosed = 4
  OpenPORequested = 5
  OpenPORequestExecuted = 6
  MaterialReceiptAtSite = 7
  POVerification = 8
  POApproval = 9
  TCPOIssued = 10
End Enum
Public Enum pakPkgStates
  Free = 1
  ReadyToDespatch = 2
  UnderReceiveAtSite = 3
  ReceivedAtSite = 4
  UnderReceiveAtPort = 5
  ReceivedAtPort = 6
  DespatchedFromPort = 7
End Enum
Public Enum pakErpPOTypes
  ISGECEngineered = 1
  Package = 2
  Boughtout = 3
End Enum
Public Enum siteIssueTypes
  Issue = 1
  Scrap = 2
  Lost = 3
End Enum
Public Enum siteIssueStates
  Free = 1
  UnderIssue = 2
  Issued = 3
  Returned = 4
End Enum
Public Enum siteInventoryStates
  Free = 1
  PackagePending = 2
  Received = 3
  Closed = 4
End Enum
Public Enum siteReceiveStates
  Free = 1
  PackagePending = 2
  Received = 3
End Enum
Public Enum pakQCStates
  Returned = 1
  Free = 2
  UnderQualityInspection = 3
  QCCompleted = 4
  PackingListCreated = 5
End Enum