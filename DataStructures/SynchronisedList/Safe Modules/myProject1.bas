Attribute VB_Name = "myProject1"
Option Explicit
Private Type codeItem
    extension As String
    module_name As String
    code_content() As String
End Type

Private Const TypeBinary As Long = 1
Private Const vbext_pp_none As Long = 0
Private Const ForReading As Long = 1, ForWriting As Long = 2, ForAppending As Long = 8

Private Function getCodeDefinition(ByVal itemNo As Long) As codeItem
    With getCodeDefinition
        Select Case itemNo
            Case 1
                .extension = ".cls"
                .module_name = "dummyRange"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiAxLjAgQ0xBU1MNCkJFR0lODQogIE11bHRpVXNlID0gLTEgICdUcnVlDQpFTkQNCkF0dHJpYnV0ZSBWQl9OYW1lID0gImR1bW15UmFuZ2UiDQpBdHRyaWJ1dGUgVkJfR2xvYmFsTmFtZVNwYWNlID0gRmFsc2UNCkF0dHJpYnV0ZSBWQl9DcmVhdGFibGUgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX1ByZWRlY2xhcmVkSWQgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX0V4cG9zZWQgPSBGYWxzZQ0KJ0BGb2xkZXIoIkNvZGVSZXZpZXciKQ0KT3B0aW9uIEV4cGxpY2l0DQoNClByaXZhdGUgY2VsbFZhbHMgQXMgT2JqZWN0ICdjb250YWlucyBkdW1teSBjZWxsIGRhdGENCg0KUHVibGljIFByb3BlcnR5IEdldCBDZWxscyhCeVZhbCByb3dOdW0gQXMgTG9uZywgQnlWYWwgY29sTnVtIEFzIExvbmcpIEFzIFN0cmluZw0KICAgIENlbGxzID0gY2VsbFZhbHMuaXRlbShnZXRLZXkocm93TnVtLCBjb2xOdW0pKQ0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBMZXQgQ2VsbHMoQnlWYWwgcm93TnVtIEFzIExvbmcsIEJ5VmFsIGNvbE51bSBBcyBMb25nLCBCeVZhbCBuZXdWYWwgQXMgU3RyaW5nKQ0KICAgIGNlbGxWYWxzLml0ZW0oZ2V0S2V5KHJvd051bSwgY29sTnVtKSkgPSBuZXdWYWwNCkVuZCBQcm9wZXJ0eQ0KDQpQcml2YXRlIFN1YiBDbGFzc19Jbml0aWFsaXplKCkNCiAgICBTZXQg" & _
"Y2VsbFZhbHMgPSBDcmVhdGVPYmplY3QoIlNjcmlwdGluZy5EaWN0aW9uYXJ5IikNCkVuZCBTdWINClByaXZhdGUgRnVuY3Rpb24gZ2V0S2V5KEJ5VmFsIHIgQXMgTG9uZywgQnlWYWwgYyBBcyBMb25nKSBBcyBTdHJpbmcNCiAgICBnZXRLZXkgPSAiaXRlbSIgJiByICYgIl8iICYgYw0KRW5kIEZ1bmN0aW9uDQo="
            Case 2
                .extension = ".cls"
                .module_name = "FormRunner"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiAxLjAgQ0xBU1MNCkJFR0lODQogIE11bHRpVXNlID0gLTEgICdUcnVlDQpFTkQNCkF0dHJpYnV0ZSBWQl9OYW1lID0gIkZvcm1SdW5uZXIiDQpBdHRyaWJ1dGUgVkJfR2xvYmFsTmFtZVNwYWNlID0gRmFsc2UNCkF0dHJpYnV0ZSBWQl9DcmVhdGFibGUgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX1ByZWRlY2xhcmVkSWQgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX0V4cG9zZWQgPSBGYWxzZQ0KJ0BGb2xkZXIoIkNvZGVSZXZpZXciKQ0KT3B0aW9uIEV4cGxpY2l0DQoNClByaXZhdGUgV2l0aEV2ZW50cyB1c2VySW50ZXJmYWNlIEFzIEV4YW1wbGVGb3JtDQpBdHRyaWJ1dGUgdXNlckludGVyZmFjZS5WQl9WYXJIZWxwSUQgPSAtMQ0KUHJpdmF0ZSBXaXRoRXZlbnRzIHN5bmNocm8gQXMgU3luY2hyb25pc2VkTGlzdA0KQXR0cmlidXRlIHN5bmNocm8uVkJfVmFySGVscElEID0gLTENCg0KUHJpdmF0ZSBUeXBlIHRSdW5uZXINCiAgICBoZWFkZXJSYW5nZSBBcyBSYW5nZQ0KICAgIFVJRm9ybSBBcyBFeGFtcGxlRm9ybQ0KICAgIGRhdGEgQXMgU3luY2hyb25pc2VkTGlzdA0KICAgIFNvcnRDb21wYXJlciBBcyBDYWxsQnlOYW1lQ29tcGFyZXINCiAgICBGaWx0ZXJDb21wYXJlciBBcyBDYWxsQnlOYW1lQ29tcGFyZXINCiAgICBmaWx0ZXJPYmogQXMgZHVtbXlSYW5nZQ0KRW5kIFR5cGUNClBy" & _
"aXZhdGUgdGhpcyBBcyB0UnVubmVyDQoNCg0KUHJpdmF0ZSBTdWIgQ2xhc3NfSW5pdGlhbGl6ZSgpDQoNCiAgICBTZXQgc3luY2hybyA9IE5ldyBTeW5jaHJvbmlzZWRMaXN0DQogICAgU2V0IHRoaXMuZGF0YSA9IHN5bmNocm8NCiAgICBTZXQgdGhpcy5Tb3J0Q29tcGFyZXIgPSBOZXcgQ2FsbEJ5TmFtZUNvbXBhcmVyDQogICAgU2V0IHRoaXMuRmlsdGVyQ29tcGFyZXIgPSBOZXcgQ2FsbEJ5TmFtZUNvbXBhcmVyDQogICAgU2V0IHRoaXMuZmlsdGVyT2JqID0gTmV3IGR1bW15UmFuZ2UNCkVuZCBTdWINCg0KUHVibGljIFN1YiBpbml0KEJ5VmFsIGRhdGFUYWJsZSBBcyBMaXN0T2JqZWN0KQ0KICAgIFNldCB0aGlzLmhlYWRlclJhbmdlID0gZGF0YVRhYmxlLkhlYWRlclJvd1JhbmdlDQogICAgDQogICAgU2V0IHVzZXJJbnRlcmZhY2UgPSBOZXcgRXhhbXBsZUZvcm0NCiAgICBTZXQgdGhpcy5VSUZvcm0gPSB1c2VySW50ZXJmYWNlDQogICAgdGhpcy5VSUZvcm0uZGF0YURpc3BsYXlCb3guQ29sdW1uQ291bnQgPSB0aGlzLmhlYWRlclJhbmdlLkNlbGxzLkNvdW50ICdzZXQgbnVtYmVyIG9mIGNvbHVtbnMNCiAgICANCiAgICB0aGlzLlVJRm9ybS5wb3B1bGF0ZUZpbHRlckJveCB0aGlzLmhlYWRlclJhbmdlDQogICAgdGhpcy5VSUZvcm0ucG9wdWxhdGVTb3J0Qm94IHRoaXMuaGVhZGVyUmFu" & _
"Z2UNCiAgICANCiAgICAnc2hvdyBmb3JtIGFuZCBzdGFydCBhZGRpbmcgZGF0YSB0byBpdA0KICAgIHRoaXMuVUlGb3JtLlNob3cgRmFsc2UNCiAgICBzeW5jaHJvLkFkZCBkYXRhVGFibGUuRGF0YUJvZHlSYW5nZS5Sb3dzDQogICAgDQoNCkVuZCBTdWINCg0KDQpQcml2YXRlIFN1YiBzeW5jaHJvX09yZGVyQ2hhbmdlZChCeVZhbCBmaXJzdENoYW5nZUluZGV4IEFzIExvbmcpDQogICAgdGhpcy5VSUZvcm0uQ2xlYXJGcm9tSW5kZXggZmlyc3RDaGFuZ2VJbmRleCAnbGlzdGJveCBpcyAwIGluZGV4ZWQgdG9vDQogICAgRGltIGkgQXMgTG9uZw0KICAgIEZvciBpID0gZmlyc3RDaGFuZ2VJbmRleCBUbyB0aGlzLmRhdGEuQ29udGVudERhdGEuQ291bnQgLSAxICcwIGluZGV4ZWQNCiAgICAgICAgdGhpcy5VSUZvcm0uQWRkSXRlbSB0aGlzLmRhdGEuQ29udGVudERhdGEoaSkNCiAgICBOZXh0DQogICAgdGhpcy5VSUZvcm0uUmVwYWludA0KICAgIERvRXZlbnRzDQpFbmQgU3ViDQoNClByaXZhdGUgU3ViIHVzZXJJbnRlcmZhY2VfZmlsdGVyTW9kZVNldChCeVZhbCBGaWx0ZXJCeSBBcyBTdHJpbmcsIEJ5VmFsIEZpbHRlclZhbHVlIEFzIFN0cmluZykNCiAgICB0aGlzLkZpbHRlckNvbXBhcmVyLmluaXQgIkNlbGxzIiwgVmJHZXQsIDEsIGNvbHVtbkluZGV4RnJvbU5hbWUoRmlsdGVyQnkpDQog" & _
"ICAgdGhpcy5maWx0ZXJPYmouQ2VsbHMoMSwgY29sdW1uSW5kZXhGcm9tTmFtZShGaWx0ZXJCeSkpID0gRmlsdGVyVmFsdWUNCiAgICB0aGlzLmRhdGEuRmlsdGVyIHRoaXMuZmlsdGVyT2JqLCB0aGlzLkZpbHRlckNvbXBhcmVyDQpFbmQgU3ViDQoNClByaXZhdGUgU3ViIHVzZXJJbnRlcmZhY2Vfc29ydE1vZGVTZXQoQnlWYWwgU29ydEJ5IEFzIFN0cmluZykNCiAgICB0aGlzLlNvcnRDb21wYXJlci5pbml0ICJDZWxscyIsIFZiR2V0LCAxLCBjb2x1bW5JbmRleEZyb21OYW1lKFNvcnRCeSkNCiAgICB0aGlzLmRhdGEuU29ydCB0aGlzLlNvcnRDb21wYXJlcg0KRW5kIFN1Yg0KUHJpdmF0ZSBGdW5jdGlvbiBjb2x1bW5JbmRleEZyb21OYW1lKEJ5VmFsIGNvbE5hbWUgQXMgU3RyaW5nKSBBcyBMb25nDQogICAgQ29uc3QgRVhBQ1RfTUFUQ0ggQXMgTG9uZyA9IDANCiAgICBEaW0gcmVzdWx0DQogICAgT24gRXJyb3IgUmVzdW1lIE5leHQNCiAgICByZXN1bHQgPSBXb3Jrc2hlZXRGdW5jdGlvbi5NYXRjaChjb2xOYW1lLCB0aGlzLmhlYWRlclJhbmdlLCBFWEFDVF9NQVRDSCkNCiAgICBjb2x1bW5JbmRleEZyb21OYW1lID0gSUlmKEVyci5OdW1iZXIgPSAwLCByZXN1bHQsIDEpDQpFbmQgRnVuY3Rpb24NCg=="
            Case 3
                .extension = ".frm"
                .module_name = "ExampleForm"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiA1LjAwDQpCZWdpbiB7QzYyQTY5RjAtMTZEQy0xMUNFLTlFOTgtMDBBQTAwNTc0QTRGfSBFeGFtcGxlRm9ybSANCiAgIENhcHRpb24gICAgICAgICA9ICAgIkV4YW1wbGUgQXBwIg0KICAgQ2xpZW50SGVpZ2h0ICAgID0gICA2MjEwDQogICBDbGllbnRMZWZ0ICAgICAgPSAgIDEyMA0KICAgQ2xpZW50VG9wICAgICAgID0gICA0NjUNCiAgIENsaWVudFdpZHRoICAgICA9ICAgNzk2NQ0KICAgT2xlT2JqZWN0QmxvYiAgID0gICAiRXhhbXBsZUZvcm0uZnJ4IjowMDAwDQogICBTdGFydFVwUG9zaXRpb24gPSAgIDEgICdDZW50ZXJPd25lcg0KRW5kDQpBdHRyaWJ1dGUgVkJfTmFtZSA9ICJFeGFtcGxlRm9ybSINCkF0dHJpYnV0ZSBWQl9HbG9iYWxOYW1lU3BhY2UgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX0NyZWF0YWJsZSA9IEZhbHNlDQpBdHRyaWJ1dGUgVkJfUHJlZGVjbGFyZWRJZCA9IFRydWUNCkF0dHJpYnV0ZSBWQl9FeHBvc2VkID0gRmFsc2UNCidARm9sZGVyKCJDb2RlUmV2aWV3IikNCk9wdGlvbiBFeHBsaWNpdA0KDQpQdWJsaWMgRXZlbnQgc29ydE1vZGVTZXQoQnlWYWwgU29ydEJ5IEFzIFN0cmluZykNClB1YmxpYyBFdmVudCBmaWx0ZXJNb2RlU2V0KEJ5VmFsIEZpbHRlckJ5IEFzIFN0cmluZywgQnlWYWwgRmlsdGVyVmFsdWUgQXMgU3RyaW5nKQ0KDQonR1VJDQpQcml2" & _
"YXRlIGJvb2xFbnRlciBBcyBCb29sZWFuDQoNCg0KJw0KJ1B1YmxpYyBPcmRlcnNMaXN0IEFzIG1zY29ybGliLkFycmF5TGlzdA0KJ1ByaXZhdGUgcGMgQXMgcHJvcGVydHlDb21wYXJlcg0KJw0KJ1ByaXZhdGUgU3ViIFVzZXJGb3JtX0luaXRpYWxpemUoKQ0KJyAgICBEaW0gY2VsbCBBcyBSYW5nZQ0KJyAgICBTZXQgT3JkZXJzTGlzdCA9IE5ldyBBcnJheUxpc3QNCicgICAgU2V0IHBjID0gTmV3IHByb3BlcnR5Q29tcGFyZXINCicNCicgICAgV2l0aCBXb3Jrc2hlZXRzKCJPcmRlcnMiKQ0KJyAgICAgICAgRm9yIEVhY2ggY2VsbCBJbiAuUmFuZ2UoIkEyIiwgLlJhbmdlKCJBIiAmIC5Sb3dzLkNvdW50KS5FbmQoeGxVcCkpDQonICAgICAgICAgICAgT3JkZXJzTGlzdC5BZGQgY2VsbC5SZXNpemUoMSwgOCkNCicgICAgICAgIE5leHQNCicNCicgICAgICAgIEZvciBFYWNoIGNlbGwgSW4gLlJhbmdlKCJBMSIpLlJlc2l6ZSgxLCA4KQ0KJyAgICAgICAgICAgIGNib1NvcnRCeS5BZGRJdGVtIGNlbGwuVmFsdWUNCicgICAgICAgIE5leHQNCicNCicgICAgRW5kIFdpdGgNCicNCicgICAgY2JvU29ydEJ5LkFkZEl0ZW0gIlJvdyINCicNCicgICAgRmlsbE9yZGVyc0xpc3RCb3gNCidFbmQgU3ViDQonDQoNCicNCidQcml2YXRlIFN1YiBidG5SZXZlcnNlX0NsaWNrKCkNCicgICAgT3JkZXJzTGlz" & _
"dC5SZXZlcnNlDQonICAgIEZpbGxPcmRlcnNMaXN0Qm94DQonRW5kIFN1Yg0KJw0KJ1ByaXZhdGUgU3ViIGNib1NvcnRCeV9DaGFuZ2UoKQ0KJyAgICBJZiBjYm9Tb3J0QnkuTGlzdEluZGV4ID0gLTEgVGhlbiBFeGl0IFN1Yg0KJw0KJyAgICBTZWxlY3QgQ2FzZSBjYm9Tb3J0QnkuTGlzdEluZGV4DQonICAgICAgICBDYXNlIElzIDwgOA0KJyAgICAgICAgICAgIHBjLkluaXQgIkNlbGxzIiwgVmJHZXQsIDEsIGNib1NvcnRCeS5MaXN0SW5kZXggKyAxDQonICAgICAgICBDYXNlIDgNCicgICAgICAgICAgICBwYy5Jbml0ICJSb3ciLCBWYkdldA0KJyAgICBFbmQgU2VsZWN0DQonDQonICAgIE9yZGVyc0xpc3QuU29ydF8yIHBjDQonICAgIEZpbGxPcmRlcnNMaXN0Qm94DQonRW5kIFN1Yg0KDQonRm9ybQ0KDQoNCidGb3JtIENvbnRyb2wgTWV0aG9kcw0KDQoNClN1YiBwb3B1bGF0ZVNvcnRCb3goQnlWYWwgb3B0aW9ucyBBcyBWYXJpYW50KQ0KICAgIE1lLlNvcnRCeS5MaXN0ID0gZG91YmxlVHJhbnNwb3NlKG9wdGlvbnMpDQpFbmQgU3ViDQogICAgDQogICAgDQpTdWIgcG9wdWxhdGVGaWx0ZXJCb3goQnlWYWwgb3B0aW9ucyBBcyBWYXJpYW50KQ0KICAgIE1lLkZpbHRlckJ5Lkxpc3QgPSBkb3VibGVUcmFuc3Bvc2Uob3B0aW9ucykNCkVuZCBTdWINCg0KDQpQdWJsaWMgU3ViIERpc3BsYXlE" & _
"YXRhKEJ5UmVmIGRhdGFBcnJheSBBcyBWYXJpYW50KQ0KICAgIElmIElzQXJyYXkoZGF0YUFycmF5KSBBbmQgQXJyYXlTdXBwb3J0Lk51bWJlck9mQXJyYXlEaW1lbnNpb25zKGRhdGFBcnJheSkgPSAxIFRoZW4NCiAgICAgICAgZGF0YURpc3BsYXlCb3guTGlzdCA9IGRvdWJsZVRyYW5zcG9zZShkYXRhQXJyYXkpDQogICAgRWxzZQ0KICAgICAgICBFcnIuUmFpc2UgNQ0KICAgIEVuZCBJZg0KRW5kIFN1Yg0KDQpQdWJsaWMgU3ViIFJlbW92ZUl0ZW0oQnlWYWwgaXRlbUluZGV4IEFzIExvbmcpDQogICAgZGF0YURpc3BsYXlCb3guUmVtb3ZlSXRlbSBpdGVtSW5kZXgNCkVuZCBTdWINCg0KUHVibGljIFN1YiBBZGRJdGVtKGl0ZW1BcnJheSBBcyBWYXJpYW50KQ0KDQogICAgSWYgSXNBcnJheShpdGVtQXJyYXkpIFRoZW4gJ2Fzc3VtZSAxIGluZGV4ZWQNCiAgICAgICAgRGltIHRyYW5zcG9zZWRBcnJheQ0KICAgICAgICB0cmFuc3Bvc2VkQXJyYXkgPSBkb3VibGVUcmFuc3Bvc2UoaXRlbUFycmF5KQ0KICAgICAgICBXaXRoIGRhdGFEaXNwbGF5Qm94DQogICAgICAgICAgICAuQWRkSXRlbQ0KICAgICAgICAgICAgRGltIGkgQXMgTG9uZw0KICAgICAgICAgICAgRm9yIGkgPSAwIFRvIC5Db2x1bW5Db3VudCAtIDENCiAgICAgICAgICAgICAgICAuTGlzdCgubGlzdENvdW50IC0gMSwgaSkgPSB0" & _
"cmFuc3Bvc2VkQXJyYXkoaSArIDEpDQogICAgICAgICAgICBOZXh0DQogICAgICAgIEVuZCBXaXRoDQogICAgRWxzZQ0KICAgICAgICBFcnIuUmFpc2UgNQ0KICAgIEVuZCBJZg0KRW5kIFN1Yg0KDQpQdWJsaWMgU3ViIENsZWFyRnJvbUluZGV4KHN0YXJ0aW5nSW5kZXggQXMgTG9uZykNCiAgICBEaW0gaSBBcyBMb25nDQogICAgRGltIGxpc3RDb3VudCBBcyBMb25nDQogICAgbGlzdENvdW50ID0gZGF0YURpc3BsYXlCb3gubGlzdENvdW50DQogICAgJ25vdGhpbmcgdG8gY2xlYXIgaWYgZmlyc3QgY2hhbmdlID4gZW5kIG9mIGxpc3QgMCBpbmRleGVkDQogICAgSWYgbGlzdENvdW50ID0gc3RhcnRpbmdJbmRleCBUaGVuIEV4aXQgU3ViDQogICAgRm9yIGkgPSBsaXN0Q291bnQgLSAxIFRvIHN0YXJ0aW5nSW5kZXggU3RlcCAtMSAnY291bnQgYmFja3dhcmRzDQogICAgICAgIFJlbW92ZUl0ZW0gaQ0KICAgIE5leHQNCkVuZCBTdWINCg0KUHJpdmF0ZSBGdW5jdGlvbiBkb3VibGVUcmFuc3Bvc2UoQnlWYWwgYXJyYXlUb1RyYW5zcG9zZSBBcyBWYXJpYW50KSBBcyBWYXJpYW50DQogICAgZG91YmxlVHJhbnNwb3NlID0gV29ya3NoZWV0RnVuY3Rpb24uVHJhbnNwb3NlKFdvcmtzaGVldEZ1bmN0aW9uLlRyYW5zcG9zZShhcnJheVRvVHJhbnNwb3NlKSkNCkVuZCBGdW5jdGlvbg0KDQonRm9ybSBH" & _
"VUkNCg0KUHJpdmF0ZSBTdWIgRmlsdGVyVmFsdWVfRW50ZXIoKQ0KICAgIGJvb2xFbnRlciA9IFRydWUNCkVuZCBTdWINCg0KUHJpdmF0ZSBTdWIgRmlsdGVyVmFsdWVfTW91c2VEb3duKEJ5VmFsIEJ1dHRvbiBBcyBJbnRlZ2VyLCBCeVZhbCBTaGlmdCBBcyBJbnRlZ2VyLCBfDQpCeVZhbCBYIEFzIFNpbmdsZSwgQnlWYWwgWSBBcyBTaW5nbGUpDQogICAgSWYgYm9vbEVudGVyID0gVHJ1ZSBUaGVuDQogICAgICAgIFdpdGggRmlsdGVyVmFsdWUNCiAgICAgICAgICAgIC5TZWxTdGFydCA9IDANCiAgICAgICAgICAgIC5TZWxMZW5ndGggPSBMZW4oLlRleHQpDQogICAgICAgIEVuZCBXaXRoDQogICAgICAgIGJvb2xFbnRlciA9IEZhbHNlDQogICAgRW5kIElmDQpFbmQgU3ViDQoNClByaXZhdGUgU3ViIFNvcnRCdXR0b25fQ2xpY2soKQ0KICAgIFJhaXNlRXZlbnQgc29ydE1vZGVTZXQoTWUuU29ydEJ5LlZhbHVlKQ0KRW5kIFN1Yg0KDQpQcml2YXRlIFN1YiBGaWx0ZXJCdXR0b25fQ2xpY2soKQ0KICAgIFJhaXNlRXZlbnQgZmlsdGVyTW9kZVNldChNZS5GaWx0ZXJCeS5WYWx1ZSwgTWUuRmlsdGVyVmFsdWUuVmFsdWUpDQpFbmQgU3ViDQo="
            Case 4
                .extension = ".bas"
                .module_name = "CodeReviewTest"
                ReDim .code_content(0 To 0)
                .code_content(0) = "QXR0cmlidXRlIFZCX05hbWUgPSAiQ29kZVJldmlld1Rlc3QiDQonQEZvbGRlcigiQ29kZVJldmlldyIpDQpPcHRpb24gRXhwbGljaXQNCg0KU3ViIHNob3dGb3JtKCkNCiAgICBTdGF0aWMgcnVubmVyIEFzIE5ldyBGb3JtUnVubmVyDQogICAgcnVubmVyLmluaXQgV29ya3NoZWV0cygiZGF0YSIpLkxpc3RPYmplY3RzKCJFeGFtcGxlRGF0YSIpDQpFbmQgU3ViDQoNClN1YiBjb21waWxlU3luY2hyb2xpc3QoKQ0KVG9vbGJveC5Db21wcmVzc1Byb2plY3QgVGhpc1dvcmtib29rLCAiRmlsdGVybGlzdCIgXw0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgLCAiRmlsdGVybGlzdFV0aWxzIiBfDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAsICJBcnJheVN1cHBvcnQiIF8NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICwgIkZpbHRlclJ1bm5lciIgXw0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgLCAiU3luY2hyb0xpc3RVdGlscyIgXw0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgLCAiQ29udGVudERhdGFXcmFwcGVyIiBfDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAsICJMaXN0QnVmZmVyIiBfDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAsICJT" & _
"b3VyY2VEYXRhV3JhcHBlciIgXw0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgLCAiU3luY2hyb25pc2VkTGlzdCINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIA0KVG9vbGJveC5Db21wcmVzc1Byb2plY3QgVGhpc1dvcmtib29rLCAiQ29kZVJldmlld1Rlc3QiIF8NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICwgIkV4YW1wbGVGb3JtIiBfDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAsICJGb3JtUnVubmVyIiBfDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAsICJkdW1teVJhbmdlIg0KRW5kIFN1Yg0KDQo="
        Case Else
            .extension = "missing"
        End Select
    End With
End Function

Public Sub Extract()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim code_module As codeItem
    Dim savedPath As String, basePath As String
    Dim i As Long
   'check if vbproject accessible
    If Not ProjectAccessible(wb) Then
        MsgBox "The VBA project cannot be accessed programmatically"
        Exit Sub
    End If
   'check if temp folder acessible
    i = 0
    basePath = Environ$("Temp") & "\"
    Do While True
        i = i + 1
        code_module = getCodeDefinition(i)
        If code_module.extension = "missing" Then
            Exit Do
        Else
            savedPath = createFile(code_module, basePath)
            importFile savedPath, wb
            Kill savedPath
        End If
    Loop
    RemoveModule "myProject1", wb
End Sub

Private Function ProjectAccessible(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    With wb.VBProject
        ProjectAccessible = .Protection = vbext_pp_none
        ProjectAccessible = ProjectAccessible And Err.Number = 0
    End With
End Function

Private Function createFile(ByRef definition As codeItem, ByVal filePath As String) As String
    Dim newFileObj As Object
    Set newFileObj = CreateObject("ADODB.Stream")
    newFileObj.Type = TypeBinary
   'Open the stream and write binary data
    newFileObj.Open
   'create file from x64 string
    With definition
        Dim bytes() As Byte
        Dim fullPath As String
        fullPath = filePath & .module_name & .extension
        bytes = FromBase64(Join(.code_content))
        newFileObj.Write bytes
        newFileObj.SaveToFile fullPath, ForWriting
        createFile = fullPath
    End With
End Function

Private Sub importFile(ByVal filePath As String, ByRef wb As Workbook)
    wb.VBProject.VBComponents.Import filePath
End Sub

Private Function RemoveModule(ByVal moduleName As String, ByRef book As Workbook) As Boolean
    On Error Resume Next
    With book.VBProject.VBComponents
        .Remove .item(moduleName)
    End With
    RemoveModule = Not (Err.Number = 9)
End Function

Private Function FromBase64(ByVal Text As String) As Byte()
    Dim Out() As Byte
    Dim b64(0 To 255) As Byte, str() As Byte, i As Long, j As Long, v As Long, b0 As Long, b1 As Long, b2 As Long, b3 As Long
    Out = vbNullString
    If Len(Text) Then Else Exit Function

    str = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For i = 2 To UBound(str) Step 2
        b64(str(i)) = i \ 2
    Next

    ReDim Out(0 To ((Len(Text) + 3) \ 4) * 3 - 1)
    str = Text & String$(2, 0)

    For i = 0 To UBound(str) - 7 Step 2
        b0 = b64(str(i))

        If b0 Then
            b1 = b64(str(i + 2))
            b2 = b64(str(i + 4))
            b3 = b64(str(i + 6))
            v = b0 * 262144 + b1 * 4096& + b2 * 64& + b3 - 266305
            Out(j) = v \ 65536
            Out(j + 1) = (v \ 256&) Mod 256
            Out(j + 2) = v Mod 256
            j = j + 3
            i = i + 6
        End If
    Next

    If b2 = 0 Then
        Out(j - 3) = (v + 65) \ 65536
        j = j - 2
    ElseIf b3 = 0 Then
        Out(j - 3) = (v + 1) \ 65536
        Out(j - 2) = ((v + 1) \ 256&) Mod 256
        j = j - 1
    End If

    ReDim Preserve Out(j - 1)
    FromBase64 = Out
End Function

