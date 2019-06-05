//
// GPO
//
OUTPUT_ENTRY pGpoEntries[14] =
{
   { TYPE_GUID, L"Identifier", 0 },
   { TYPE_STRING, L"Domain", 0 },
   { TYPE_STRING, L"Name", 0 },

   { TYPE_DATE, L"CreatedTime", 0 },
   { TYPE_DATE, L"ModifiedTime", 0 },

   { TYPE_BOOL, L"FilterDataAvailable", 0 },
   { TYPE_STRING, L"FilterName", 0 },
   { TYPE_STRING, L"FilterDescription", 0 },

   { TYPE_INTEGER, L"Computer_VersionDirectory", 0 },
   { TYPE_INTEGER, L"Computer_VersionSysvol", 0 },
   { TYPE_BOOL, L"Computer_Enabled", 0 },

   { TYPE_INTEGER, L"User_VersionDirectory", 0 },
   { TYPE_INTEGER, L"User_VersionSysvol", 0 },
   { TYPE_BOOL, L"User_Enabled", 0 },
};

OUTPUT_ENTRY pExtensionEntries[1] =
{
   { TYPE_STRING, L"Name", 0 },
};

OUTPUT_ENTRY pLinksToEntries[4] =
{
   { TYPE_STRING, L"SOMName", 0 },
   { TYPE_STRING, L"SOMPath", 0 },
   { TYPE_BOOL, L"Enabled", 0 },
   { TYPE_BOOL, L"NoOverride", 0 },
};

//
// ExtRegistry
//
OUTPUT_ENTRY pRegistryPolicyEntries[5] =
{
   { TYPE_INTEGER, L"id_policy", 0 },
   { TYPE_STRING, L"Name", 0 },
   { TYPE_STRING, L"Comment", 0 },
   { TYPE_STRING, L"State", 0 },
   { TYPE_STRING, L"Category", 0 }
};

OUTPUT_ENTRY pRegistryValuesEntries[5] =
{
   { TYPE_INTEGER, L"id_policy", 0 },
   { TYPE_STRING, L"Type", 0 },
   { TYPE_STRING, L"Name", 0 },
   { TYPE_STRING, L"State", 0 },
   { TYPE_STRING, L"Value", 0 }
};

//
// ExtSecurity
//
OUTPUT_ENTRY pSecurityOptionEntries[4] =
{
   { TYPE_STRING, L"Name", 0 },
   { TYPE_STRING, L"Value", 0 },
   { TYPE_STRING, L"DisplayName", 0 },
   { TYPE_STRING, L"DisplayString", 0 }
};

OUTPUT_ENTRY pSecuritRestrictedGroupsEntries[4] =
{
   { TYPE_STRING, L"GroupNameSID", 0 },
   { TYPE_STRING, L"GroupNameName", 0 },
   { TYPE_STRING, L"MemberOfSID", 0 },
   { TYPE_STRING, L"MemberOfName", 0 }
};

OUTPUT_ENTRY pSecurityUserRightsAssignmentEntries[3] =
{
   { TYPE_STRING, L"PrivName", 0 },
   { TYPE_STRING, L"SID", 0 },
   { TYPE_STRING, L"Name", 0 }
};

//
// ExtAppControl
//
OUTPUT_ENTRY pAppControlRuleCollectionEntries[3] =
{
   { TYPE_INTEGER, L"id_rules", 0 },
   { TYPE_STRING, L"Type", 0 },
   { TYPE_STRING, L"EnforcementMode", 0 }
};

OUTPUT_ENTRY pAppControlRuleEntries[7] =
{
   { TYPE_INTEGER, L"id_rules", 0 },
   { TYPE_INTEGER, L"id_rule", 0 },
   { TYPE_STRING, L"Type", 0 },
   { TYPE_STRING, L"Name", 0 },
   { TYPE_STRING, L"Description", 0 },
   { TYPE_STRING, L"UserOrGroupSid", 0 },
   { TYPE_STRING, L"Action", 0 }
};

OUTPUT_ENTRY pAppControlConditionsEntries[4] =
{
   { TYPE_INTEGER, L"id_rules", 0 },
   { TYPE_INTEGER, L"id_rule", 0 },
   { TYPE_INTEGER, L"id_condition", 0 },
   { TYPE_STRING, L"Value", 0 }
};

//
// ExtScript
//
OUTPUT_ENTRY pScriptOptionEntries[4] =
{
   { TYPE_STRING, L"Command", 0 },
   { TYPE_STRING, L"Parameters", 0 },
   { TYPE_STRING, L"Type", 0 },
   { TYPE_STRING, L"Order", 0 }
};

//
// ExtFirewall
//
OUTPUT_ENTRY pFirewallProfileEntries[12] =
{
   { TYPE_STRING, L"Profile", 0 },
   { TYPE_STRING_TO_BOOL, L"EnableFirewall", 0 },
   { TYPE_STRING_TO_BOOL, L"DefaultInboundAction", 0 },
   { TYPE_STRING_TO_BOOL, L"DefaultOutboundAction", 0 },
   { TYPE_STRING_TO_BOOL, L"DoNotAllowExceptions", 0 },
   { TYPE_STRING_TO_BOOL, L"DisableNotifications", 0 },
   { TYPE_STRING_TO_BOOL, L"AllowLocalPolicyMerge", 0 },
   { TYPE_STRING_TO_BOOL, L"AllowLocalIPsecPolicyMerge", 0 },
   { TYPE_STRING, L"LogFilePath", 0 },
   { TYPE_STRING_TO_INTEGER, L"LogFileSize", 0 },
   { TYPE_STRING_TO_BOOL, L"LogDroppedPackets", 0 },
   { TYPE_STRING_TO_BOOL, L"LogSuccessfulConnections", 0 }
};

OUTPUT_ENTRY pFirewallRuleEntries[21] =
{
   { TYPE_STRING, L"Type", 0 },
   { TYPE_STRING, L"Profile", 0 },
   { TYPE_STRING, L"Version", 0 },
   { TYPE_STRING, L"Action", 0 },
   { TYPE_STRING, L"Name", 0 },
   { TYPE_STRING, L"Desc", 0 },
   { TYPE_STRING, L"Active", 0 },
   { TYPE_STRING, L"EmbedCtxt", 0 },
   { TYPE_STRING, L"App", 0 },
   { TYPE_STRING, L"Svc", 0 },
   { TYPE_STRING, L"Dir", 0 },
   { TYPE_STRING, L"Protocol", 0 },
   { TYPE_STRING, L"ICMP4", 0 },
   { TYPE_STRING, L"ICMP6", 0 },
   { TYPE_STRING, L"LPort", 0 },
   { TYPE_STRING, L"RPort", 0 },
   { TYPE_STRING, L"LA4", 0 },
   { TYPE_STRING, L"LA6", 0 },
   { TYPE_STRING, L"RA4", 0 },
   { TYPE_STRING, L"RA6", 0 },
   { TYPE_STRING, L"Security", 0 }
};

//
// Bases
//
OUTPUT_BASE OutputBases[] =
{
   { L"gpo", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pGpoEntries) / sizeof(OUTPUT_ENTRY), pGpoEntries, NULL },
   { L"extension", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pExtensionEntries) / sizeof(OUTPUT_ENTRY), pExtensionEntries, NULL },
   { L"linksto", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pLinksToEntries) / sizeof(OUTPUT_ENTRY), pLinksToEntries, NULL },

   { L"registry_policy", OUTPUT_OPTION_ADD_ID_GPO | OUTPUT_OPTION_ADD_SCOPE, sizeof(pRegistryPolicyEntries) / sizeof(OUTPUT_ENTRY), pRegistryPolicyEntries, NULL },
   { L"registry_value", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pRegistryValuesEntries) / sizeof(OUTPUT_ENTRY), pRegistryValuesEntries, NULL },

   { L"security_option", OUTPUT_OPTION_ADD_ID_GPO | OUTPUT_OPTION_ADD_SCOPE, sizeof(pSecurityOptionEntries)/sizeof(OUTPUT_ENTRY), pSecurityOptionEntries, NULL },
   { L"security_restricted_group", OUTPUT_OPTION_ADD_ID_GPO | OUTPUT_OPTION_ADD_SCOPE, sizeof(pSecuritRestrictedGroupsEntries) / sizeof(OUTPUT_ENTRY), pSecuritRestrictedGroupsEntries, NULL },
   { L"security_userright_assignment", OUTPUT_OPTION_ADD_ID_GPO | OUTPUT_OPTION_ADD_SCOPE, sizeof(pSecurityUserRightsAssignmentEntries) / sizeof(OUTPUT_ENTRY), pSecurityUserRightsAssignmentEntries, NULL },

   { L"appctl_rule_collection", OUTPUT_OPTION_ADD_ID_GPO | OUTPUT_OPTION_ADD_SCOPE, sizeof(pAppControlRuleCollectionEntries) / sizeof(OUTPUT_ENTRY), pAppControlRuleCollectionEntries, NULL },
   { L"appctl_rule", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pAppControlRuleEntries) / sizeof(OUTPUT_ENTRY), pAppControlRuleEntries, NULL },
   { L"appctl_condition", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pAppControlConditionsEntries) / sizeof(OUTPUT_ENTRY), pAppControlConditionsEntries, NULL },

   { L"script", OUTPUT_OPTION_ADD_ID_GPO | OUTPUT_OPTION_ADD_SCOPE, sizeof(pScriptOptionEntries) / sizeof(OUTPUT_ENTRY), pScriptOptionEntries, NULL },

   { L"firewall_profile", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pFirewallProfileEntries) / sizeof(OUTPUT_ENTRY), pFirewallProfileEntries, NULL },
   { L"firewall_rule", OUTPUT_OPTION_ADD_ID_GPO, sizeof(pFirewallRuleEntries) / sizeof(OUTPUT_ENTRY), pFirewallRuleEntries, NULL }
};
