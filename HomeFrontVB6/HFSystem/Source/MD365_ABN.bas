Attribute VB_Name = "MD365_ABN"
'-------------------------------------------------------------
' Financial Dimension Mapping
'-------------------------------------------------------------
'VALUE              SEND WITH       MAPPED TO
'legal entity       job & PO        System setting, one per HF division.
'D01-Division       job & PO        System setting, one per HF division.
'D02-Function       job             New field on tblLocality
'D03-CostCenter     job             New field on communityphase
'D06-Brand          job             New field on communityphase
'D04-SpendCat       PO              Debitacct from poindex/costcode
'
'Group costcode description will contain a coded string in the format "ProjectCategory / ItemCode". Expected values are
'  "JobConsumables_Item / Proj_IntC"
'  "JobSubcontractor_Item / Proj_IntS"
'  "JobFulfillment_Item / Proj_IntF"
'

