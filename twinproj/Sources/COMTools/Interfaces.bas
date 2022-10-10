Attribute VB_Name = "Interfaces"
'@Folder "OLE"
Option Explicit

'@EntryPoint
Public Function InterfacesDict() As KeyedCollection
    Static dict As KeyedCollection
    '
    If dict Is Nothing Then
        Set dict = New KeyedCollection
        '
        'dict.Add "", "{}"
        dict.Add "IUnknown", "{00000000-0000-0000-C000-000000000046}"
        dict.Add "IDispatch", "{00020400-0000-0000-C000-000000000046}"
        dict.Add "IControl", "{04598FC6-866C-11CF-AB7C-00AA00C08FCF}"
        dict.Add "Control", "{909E0AE0-16DC-11CE-9E98-00AA00574A4F}"
        dict.Add "ControlEvents", "{9A4BBF53-4E46-101B-8BBD-00AA003E3B29}"
        dict.Add "TextBox", "{8BD21D10-EC42-11CE-9E0D-00AA006002F3}"
        dict.Add "IMdcText", "{8BD21D13-EC42-11CE-9E0D-00AA006002F3}"
        dict.Add "MdcTextEvents", "{8BD21D12-EC42-11CE-9E0D-00AA006002F3}"
        dict.Add "UserForm", "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}"
        dict.Add "_UserForm", "{04598FC8-866C-11CF-AB7C-00AA00C08FCF}"
        dict.Add "FormEvents", "{5B9D8FC8-4A71-101B-97A6-00000B65C08B}"
        dict.Add "IOptionFrame", "{29B86A70-F52E-11CE-9BCE-00AA00608E01}"
        dict.Add "Label", "{978C9E23-D4B0-11CE-BF2D-00AA003F40D0}"
        dict.Add "ILabelControl", "{04598FC1-866C-11CF-AB7C-00AA00C08FCF}"
        dict.Add "LabelControlEvents", "{978C9E22-D4B0-11CE-BF2D-00AA003F40D0}"
        dict.Add "IConnectionPoint", "{B196B286-BAB4-101A-B69C-00AA00341D07}"
        dict.Add "IConnectionPointContainer", "{B196B284-BAB4-101A-B69C-00AA00341D07}"
        dict.Add "IPropertyNotifySink", "{9BFBBC02-EFF1-101A-84ED-00AA00341D07}"
        dict.Add "IMarshall", "{00000003-0000-0000-C000-000000000046}"
        dict.Add "_DClass", "{FCFB3D2B-A0FA-1068-A738-08002B3371B5}"
        dict.Add "ISupportErrorInfo", "{DF0B3D60-548F-101B-8E65-08002B2BD119}"
        dict.Add "IClassModuleEvt", "{FCFB3D21-A0FA-1068-A738-08002B3371B5}"
        dict.Add "Class", "{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
        dict.Add "ITypeLib", "{00020402-0000-0000-C000-000000000046}"
        dict.Add "ITypeLib2", "{00020411-0000-0000-C000-000000000046}"
        dict.Add "IVBEComponent", "{DDD557E1-D96F-11CD-9570-00AA0051E5D4}"
    End If
    Set InterfacesDict = dict
End Function
