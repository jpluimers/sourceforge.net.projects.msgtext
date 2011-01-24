/*-- $Workfile: PropValue.cpp $ ----------------------------------------
Wrapper for an SPropValue.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Session class wraps an IMAPISession.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "PropValue.h"
#include "../Utils/Heap.h"

#define PR_BODY_HTML					PROP_TAG( PT_LONG,		0x1013)

/*--------------------------------------------------------------------*/
PropValue::PropValue(LPSPropValue pSPropValue)
{
  m_pSPropValue = pSPropValue;
} /* PropValue */

/*--------------------------------------------------------------------*/
PropValue::~PropValue()
{
  MAPIFreeBuffer(m_pSPropValue);
} /* ~PropValue */

/*--------------------------------------------------------------------*/
ULONG PropValue::tag()
{
  return m_pSPropValue->ulPropTag;
} /* tag */

/*--------------------------------------------------------------------*/
char* PropValue::type()
{
  USHORT usPropType = (USHORT)(this->tag() & 0x0000FFFF);
  char* psType;
  switch (usPropType)
  {
  case PT_UNSPECIFIED: psType ="PT_UNSPECIFIED"; break;
  case PT_NULL: psType ="PT_NULL"; break;
  case PT_I2: psType ="PT_I2"; break;
  // case PT_SHORT: psType ="PT_SHORT"; break;
  case PT_I4: psType ="PT_I4"; break;
  // case PT_LONG: psType ="PT_LONG"; break;
  case PT_I8: psType ="PT_I8"; break;
  // case PT_LONGLONG: psType ="PT_LONGLONG"; break;
  case PT_R4: psType ="PT_R4"; break;
  // case PT_FLOAT: psType ="PT_FLOAT"; break;
  case PT_R8: psType ="PT_R8"; break;
  // case PT_DOUBLE: psType ="PT_DOUBLE"; break;
  case PT_CURRENCY: psType ="PT_CURRENCY"; break;
  case PT_APPTIME: psType ="PT_APPTIME"; break;
  case PT_ERROR: psType ="PT_ERROR"; break;
  case PT_BOOLEAN: psType ="PT_BOOLEAN"; break;
  case PT_OBJECT: psType ="PT_OBJECT"; break;
  case PT_STRING8: psType ="PT_STRING8"; break;
  case PT_UNICODE: psType ="PT_UNICODE"; break;
  case PT_SYSTIME: psType ="PT_SYSTIME"; break;
  case PT_CLSID: psType ="PT_CLSID"; break;
  case PT_BINARY: psType ="PT_BINARY"; break;
  default: psType ="PT_UNKNOWN"; break;
  }
  return psType;
} /* type */

/*--------------------------------------------------------------------*/
char* PropValue::id()
{
  USHORT usPropId = (USHORT)(this->tag() >> 16);
  char* psId;
  switch (usPropId)
  {
  case PR_ACKNOWLEDGEMENT_MODE >> 16: psId = "PR_ACKNOWLEDGEMENT_MODE"; break;
  case PR_ALTERNATE_RECIPIENT_ALLOWED >> 16: psId = "PR_ALTERNATE_RECIPIENT_ALLOWED"; break;
  case PR_AUTHORIZING_USERS >> 16: psId = "PR_AUTHORIZING_USERS"; break;
  case PR_AUTO_FORWARD_COMMENT >> 16: psId = "PR_AUTO_FORWARD_COMMENT"; break;
  case PR_AUTO_FORWARDED >> 16: psId = "PR_AUTO_FORWARDED"; break;
  case PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID >> 16: psId = "PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID"; break;
  case PR_CONTENT_CORRELATOR >> 16: psId = "PR_CONTENT_CORRELATOR"; break;
  case PR_CONTENT_IDENTIFIER >> 16: psId = "PR_CONTENT_IDENTIFIER"; break;
  case PR_CONTENT_LENGTH >> 16: psId = "PR_CONTENT_LENGTH"; break;
  case PR_CONTENT_RETURN_REQUESTED >> 16: psId = "PR_CONTENT_RETURN_REQUESTED"; break;
  case PR_CONVERSATION_KEY >> 16: psId = "PR_CONVERSATION_KEY"; break;
  case PR_CONVERSION_EITS >> 16: psId = "PR_CONVERSION_EITS"; break;
  case PR_CONVERSION_WITH_LOSS_PROHIBITED >> 16: psId = "PR_CONVERSION_WITH_LOSS_PROHIBITED"; break;
  case PR_CONVERTED_EITS >> 16: psId = "PR_CONVERTED_EITS"; break;
  case PR_DEFERRED_DELIVERY_TIME >> 16: psId = "PR_DEFERRED_DELIVERY_TIME"; break;
  case PR_DELIVER_TIME >> 16: psId = "PR_DELIVER_TIME"; break;
  case PR_DISCARD_REASON >> 16: psId = "PR_DISCARD_REASON"; break;
  case PR_DISCLOSURE_OF_RECIPIENTS >> 16: psId = "PR_DISCLOSURE_OF_RECIPIENTS"; break;
  case PR_DL_EXPANSION_HISTORY >> 16: psId = "PR_DL_EXPANSION_HISTORY"; break;
  case PR_DL_EXPANSION_PROHIBITED >> 16: psId = "PR_DL_EXPANSION_PROHIBITED"; break;
  case PR_EXPIRY_TIME >> 16: psId = "PR_EXPIRY_TIME"; break;
  case PR_IMPLICIT_CONVERSION_PROHIBITED >> 16: psId = "PR_IMPLICIT_CONVERSION_PROHIBITED"; break;
  case PR_IMPORTANCE >> 16: psId = "PR_IMPORTANCE"; break;
  case PR_IPM_ID >> 16: psId = "PR_IPM_ID"; break;
  case PR_LATEST_DELIVERY_TIME >> 16: psId = "PR_LATEST_DELIVERY_TIME"; break;
  case PR_MESSAGE_CLASS >> 16: psId = "PR_MESSAGE_CLASS"; break;
  case PR_MESSAGE_DELIVERY_ID >> 16: psId = "PR_MESSAGE_DELIVERY_ID"; break;
  case PR_MESSAGE_SECURITY_LABEL >> 16: psId = "PR_MESSAGE_SECURITY_LABEL"; break;
  case PR_OBSOLETED_IPMS >> 16: psId = "PR_OBSOLETED_IPMS"; break;
  case PR_ORIGINALLY_INTENDED_RECIPIENT_NAME >> 16: psId = "PR_ORIGINALLY_INTENDED_RECIPIENT_NAME"; break;
  case PR_ORIGINAL_EITS >> 16: psId = "PR_ORIGINAL_EITS"; break;
  case PR_ORIGINATOR_CERTIFICATE >> 16: psId = "PR_ORIGINATOR_CERTIFICATE"; break;
  case PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED >> 16: psId = "PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED"; break;
  case PR_ORIGINATOR_RETURN_ADDRESS >> 16: psId = "PR_ORIGINATOR_RETURN_ADDRESS"; break;
  case PR_PARENT_KEY >> 16: psId = "PR_PARENT_KEY"; break;
  case PR_PRIORITY >> 16: psId = "PR_PRIORITY"; break;
  case PR_ORIGIN_CHECK >> 16: psId = "PR_ORIGIN_CHECK"; break;
  case PR_PROOF_OF_SUBMISSION_REQUESTED >> 16: psId = "PR_PROOF_OF_SUBMISSION_REQUESTED"; break;
  case PR_READ_RECEIPT_REQUESTED >> 16: psId = "PR_READ_RECEIPT_REQUESTED"; break;
  case PR_RECEIPT_TIME >> 16: psId = "PR_RECEIPT_TIME"; break;
  case PR_RECIPIENT_REASSIGNMENT_PROHIBITED >> 16: psId = "PR_RECIPIENT_REASSIGNMENT_PROHIBITED"; break;
  case PR_REDIRECTION_HISTORY >> 16: psId = "PR_REDIRECTION_HISTORY"; break;
  case PR_RELATED_IPMS >> 16: psId = "PR_RELATED_IPMS"; break;
  case PR_ORIGINAL_SENSITIVITY >> 16: psId = "PR_ORIGINAL_SENSITIVITY"; break;
  case PR_LANGUAGES >> 16: psId = "PR_LANGUAGES"; break;
  case PR_REPLY_TIME >> 16: psId = "PR_REPLY_TIME"; break;
  case PR_REPORT_TAG >> 16: psId = "PR_REPORT_TAG"; break;
  case PR_REPORT_TIME >> 16: psId = "PR_REPORT_TIME"; break;
  case PR_RETURNED_IPM >> 16: psId = "PR_RETURNED_IPM"; break;
  case PR_SECURITY >> 16: psId = "PR_SECURITY"; break;
  case PR_INCOMPLETE_COPY >> 16: psId = "PR_INCOMPLETE_COPY"; break;
  case PR_SENSITIVITY >> 16: psId = "PR_SENSITIVITY"; break;
  case PR_SUBJECT >> 16: psId = "PR_SUBJECT"; break;
  case PR_SUBJECT_IPM >> 16: psId = "PR_SUBJECT_IPM"; break;
  case PR_CLIENT_SUBMIT_TIME >> 16: psId = "PR_CLIENT_SUBMIT_TIME"; break;
  case PR_REPORT_NAME >> 16: psId = "PR_REPORT_NAME"; break;
  case PR_SENT_REPRESENTING_SEARCH_KEY >> 16: psId = "PR_SENT_REPRESENTING_SEARCH_KEY"; break;
  case PR_X400_CONTENT_TYPE >> 16: psId = "PR_X400_CONTENT_TYPE"; break;
  case PR_SUBJECT_PREFIX >> 16: psId = "PR_SUBJECT_PREFIX"; break;
  case PR_NON_RECEIPT_REASON >> 16: psId = "PR_NON_RECEIPT_REASON"; break;
  case PR_RECEIVED_BY_ENTRYID >> 16: psId = "PR_RECEIVED_BY_ENTRYID"; break;
  case PR_RECEIVED_BY_NAME >> 16: psId = "PR_RECEIVED_BY_NAME"; break;
  case PR_SENT_REPRESENTING_ENTRYID >> 16: psId = "PR_SENT_REPRESENTING_ENTRYID"; break;
  case PR_SENT_REPRESENTING_NAME >> 16: psId = "PR_SENT_REPRESENTING_NAME"; break;
  case PR_RCVD_REPRESENTING_ENTRYID >> 16: psId = "PR_RCVD_REPRESENTING_ENTRYID"; break;
  case PR_RCVD_REPRESENTING_NAME >> 16: psId = "PR_RCVD_REPRESENTING_NAME"; break;
  case PR_REPORT_ENTRYID >> 16: psId = "PR_REPORT_ENTRYID"; break;
  case PR_READ_RECEIPT_ENTRYID >> 16: psId = "PR_READ_RECEIPT_ENTRYID"; break;
  case PR_MESSAGE_SUBMISSION_ID >> 16: psId = "PR_MESSAGE_SUBMISSION_ID"; break;
  case PR_PROVIDER_SUBMIT_TIME >> 16: psId = "PR_PROVIDER_SUBMIT_TIME"; break;
  case PR_ORIGINAL_SUBJECT >> 16: psId = "PR_ORIGINAL_SUBJECT"; break;
  case PR_DISC_VAL >> 16: psId = "PR_DISC_VAL"; break;
  case PR_ORIG_MESSAGE_CLASS >> 16: psId = "PR_ORIG_MESSAGE_CLASS"; break;
  case PR_ORIGINAL_AUTHOR_ENTRYID >> 16: psId = "PR_ORIGINAL_AUTHOR_ENTRYID"; break;
  case PR_ORIGINAL_AUTHOR_NAME >> 16: psId = "PR_ORIGINAL_AUTHOR_NAME"; break;
  case PR_ORIGINAL_SUBMIT_TIME >> 16: psId = "PR_ORIGINAL_SUBMIT_TIME"; break;
  case PR_REPLY_RECIPIENT_ENTRIES >> 16: psId = "PR_REPLY_RECIPIENT_ENTRIES"; break;
  case PR_REPLY_RECIPIENT_NAMES >> 16: psId = "PR_REPLY_RECIPIENT_NAMES"; break;
  case PR_RECEIVED_BY_SEARCH_KEY >> 16: psId = "PR_RECEIVED_BY_SEARCH_KEY"; break;
  case PR_RCVD_REPRESENTING_SEARCH_KEY >> 16: psId = "PR_RCVD_REPRESENTING_SEARCH_KEY"; break;
  case PR_READ_RECEIPT_SEARCH_KEY >> 16: psId = "PR_READ_RECEIPT_SEARCH_KEY"; break;
  case PR_REPORT_SEARCH_KEY >> 16: psId = "PR_REPORT_SEARCH_KEY"; break;
  case PR_ORIGINAL_DELIVERY_TIME >> 16: psId = "PR_ORIGINAL_DELIVERY_TIME"; break;
  case PR_ORIGINAL_AUTHOR_SEARCH_KEY >> 16: psId = "PR_ORIGINAL_AUTHOR_SEARCH_KEY"; break;
  case PR_MESSAGE_TO_ME >> 16: psId = "PR_MESSAGE_TO_ME"; break;
  case PR_MESSAGE_CC_ME >> 16: psId = "PR_MESSAGE_CC_ME"; break;
  case PR_MESSAGE_RECIP_ME >> 16: psId = "PR_MESSAGE_RECIP_ME"; break;
  case PR_ORIGINAL_SENDER_NAME >> 16: psId = "PR_ORIGINAL_SENDER_NAME"; break;
  case PR_ORIGINAL_SENDER_ENTRYID >> 16: psId = "PR_ORIGINAL_SENDER_ENTRYID"; break;
  case PR_ORIGINAL_SENDER_SEARCH_KEY >> 16: psId = "PR_ORIGINAL_SENDER_SEARCH_KEY"; break;
  case PR_ORIGINAL_SENT_REPRESENTING_NAME >> 16: psId = "PR_ORIGINAL_SENT_REPRESENTING_NAME"; break;
  case PR_ORIGINAL_SENT_REPRESENTING_ENTRYID >> 16: psId = "PR_ORIGINAL_SENT_REPRESENTING_ENTRYID"; break;
  case PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY >> 16: psId = "PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY"; break;
  case PR_START_DATE >> 16: psId = "PR_START_DATE"; break;
  case PR_END_DATE >> 16: psId = "PR_END_DATE"; break;
  case PR_OWNER_APPT_ID >> 16: psId = "PR_OWNER_APPT_ID"; break;
  case PR_RESPONSE_REQUESTED >> 16: psId = "PR_RESPONSE_REQUESTED"; break;
  case PR_SENT_REPRESENTING_ADDRTYPE >> 16: psId = "PR_SENT_REPRESENTING_ADDRTYPE"; break;
  case PR_SENT_REPRESENTING_EMAIL_ADDRESS >> 16: psId = "PR_SENT_REPRESENTING_EMAIL_ADDRESS"; break;
  case PR_ORIGINAL_SENDER_ADDRTYPE >> 16: psId = "PR_ORIGINAL_SENDER_ADDRTYPE"; break;
  case PR_ORIGINAL_SENDER_EMAIL_ADDRESS >> 16: psId = "PR_ORIGINAL_SENDER_EMAIL_ADDRESS"; break;
  case PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE >> 16: psId = "PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE"; break;
  case PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS >> 16: psId = "PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS"; break;
  case PR_CONVERSATION_TOPIC >> 16: psId = "PR_CONVERSATION_TOPIC"; break;
  case PR_CONVERSATION_INDEX >> 16: psId = "PR_CONVERSATION_INDEX"; break;
  case PR_ORIGINAL_DISPLAY_BCC >> 16: psId = "PR_ORIGINAL_DISPLAY_BCC"; break;
  case PR_ORIGINAL_DISPLAY_CC >> 16: psId = "PR_ORIGINAL_DISPLAY_CC"; break;
  case PR_ORIGINAL_DISPLAY_TO >> 16: psId = "PR_ORIGINAL_DISPLAY_TO"; break;
  case PR_RECEIVED_BY_ADDRTYPE >> 16: psId = "PR_RECEIVED_BY_ADDRTYPE"; break;
  case PR_RECEIVED_BY_EMAIL_ADDRESS >> 16: psId = "PR_RECEIVED_BY_EMAIL_ADDRESS"; break;
  case PR_RCVD_REPRESENTING_ADDRTYPE >> 16: psId = "PR_RCVD_REPRESENTING_ADDRTYPE"; break;
  case PR_RCVD_REPRESENTING_EMAIL_ADDRESS >> 16: psId = "PR_RCVD_REPRESENTING_EMAIL_ADDRESS"; break;
  case PR_ORIGINAL_AUTHOR_ADDRTYPE >> 16: psId = "PR_ORIGINAL_AUTHOR_ADDRTYPE"; break;
  case PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS >> 16: psId = "PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS"; break;
  case PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE >> 16: psId = "PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE"; break;
  case PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS >> 16: psId = "PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS"; break;
  case PR_TRANSPORT_MESSAGE_HEADERS >> 16: psId = "PR_TRANSPORT_MESSAGE_HEADERS"; break;
  case PR_DELEGATION >> 16: psId = "PR_DELEGATION"; break;
  case PR_TNEF_CORRELATION_KEY >> 16: psId = "PR_TNEF_CORRELATION_KEY"; break;
  case PR_BODY >> 16: psId = "PR_BODY"; break;
  case PR_REPORT_TEXT >> 16: psId = "PR_REPORT_TEXT"; break;
  case PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY >> 16: psId = "PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY"; break;
  case PR_REPORTING_DL_NAME >> 16: psId = "PR_REPORTING_DL_NAME"; break;
  case PR_REPORTING_MTA_CERTIFICATE >> 16: psId = "PR_REPORTING_MTA_CERTIFICATE"; break;
  case PR_RTF_SYNC_BODY_CRC >> 16: psId = "PR_RTF_SYNC_BODY_CRC"; break;
  case PR_RTF_SYNC_BODY_COUNT >> 16: psId = "PR_RTF_SYNC_BODY_COUNT"; break;
  case PR_RTF_SYNC_BODY_TAG >> 16: psId = "PR_RTF_SYNC_BODY_TAG"; break;
  case PR_RTF_COMPRESSED >> 16: psId = "PR_RTF_COMPRESSED"; break;
  case PR_RTF_SYNC_PREFIX_COUNT >> 16: psId = "PR_RTF_SYNC_PREFIX_COUNT"; break;
  case PR_RTF_SYNC_TRAILING_COUNT >> 16: psId = "PR_RTF_SYNC_TRAILING_COUNT"; break;
  case PR_ORIGINALLY_INTENDED_RECIP_ENTRYID >> 16: psId = "PR_ORIGINALLY_INTENDED_RECIP_ENTRYID"; break;
  case PR_BODY_HTML >> 16: psId = "PR_BODY_HTML"; break;
  case PR_CONTENT_INTEGRITY_CHECK >> 16: psId = "PR_CONTENT_INTEGRITY_CHECK"; break;
  case PR_EXPLICIT_CONVERSION >> 16: psId = "PR_EXPLICIT_CONVERSION"; break;
  case PR_IPM_RETURN_REQUESTED >> 16: psId = "PR_IPM_RETURN_REQUESTED"; break;
  case PR_MESSAGE_TOKEN >> 16: psId = "PR_MESSAGE_TOKEN"; break;
  case PR_NDR_REASON_CODE >> 16: psId = "PR_NDR_REASON_CODE"; break;
  case PR_NDR_DIAG_CODE >> 16: psId = "PR_NDR_DIAG_CODE"; break;
  case PR_NON_RECEIPT_NOTIFICATION_REQUESTED >> 16: psId = "PR_NON_RECEIPT_NOTIFICATION_REQUESTED"; break;
  case PR_DELIVERY_POINT >> 16: psId = "PR_DELIVERY_POINT"; break;
  case PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED >> 16: psId = "PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED"; break;
  case PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT >> 16: psId = "PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT"; break;
  case PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY >> 16: psId = "PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY"; break;
  case PR_PHYSICAL_DELIVERY_MODE >> 16: psId = "PR_PHYSICAL_DELIVERY_MODE"; break;
  case PR_PHYSICAL_DELIVERY_REPORT_REQUEST >> 16: psId = "PR_PHYSICAL_DELIVERY_REPORT_REQUEST"; break;
  case PR_PHYSICAL_FORWARDING_ADDRESS >> 16: psId = "PR_PHYSICAL_FORWARDING_ADDRESS"; break;
  case PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED >> 16: psId = "PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED"; break;
  case PR_PHYSICAL_FORWARDING_PROHIBITED >> 16: psId = "PR_PHYSICAL_FORWARDING_PROHIBITED"; break;
  case PR_PHYSICAL_RENDITION_ATTRIBUTES >> 16: psId = "PR_PHYSICAL_RENDITION_ATTRIBUTES"; break;
  case PR_PROOF_OF_DELIVERY >> 16: psId = "PR_PROOF_OF_DELIVERY"; break;
  case PR_PROOF_OF_DELIVERY_REQUESTED >> 16: psId = "PR_PROOF_OF_DELIVERY_REQUESTED"; break;
  case PR_RECIPIENT_CERTIFICATE >> 16: psId = "PR_RECIPIENT_CERTIFICATE"; break;
  case PR_RECIPIENT_NUMBER_FOR_ADVICE >> 16: psId = "PR_RECIPIENT_NUMBER_FOR_ADVICE"; break;
  case PR_RECIPIENT_TYPE >> 16: psId = "PR_RECIPIENT_TYPE"; break;
  case PR_REGISTERED_MAIL_TYPE >> 16: psId = "PR_REGISTERED_MAIL_TYPE"; break;
  case PR_REPLY_REQUESTED >> 16: psId = "PR_REPLY_REQUESTED"; break;
  case PR_REQUESTED_DELIVERY_METHOD >> 16: psId = "PR_REQUESTED_DELIVERY_METHOD"; break;
  case PR_SENDER_ENTRYID >> 16: psId = "PR_SENDER_ENTRYID"; break;
  case PR_SENDER_NAME >> 16: psId = "PR_SENDER_NAME"; break;
  case PR_SUPPLEMENTARY_INFO >> 16: psId = "PR_SUPPLEMENTARY_INFO"; break;
  case PR_TYPE_OF_MTS_USER >> 16: psId = "PR_TYPE_OF_MTS_USER"; break;
  case PR_SENDER_SEARCH_KEY >> 16: psId = "PR_SENDER_SEARCH_KEY"; break;
  case PR_SENDER_ADDRTYPE >> 16: psId = "PR_SENDER_ADDRTYPE"; break;
  case PR_SENDER_EMAIL_ADDRESS >> 16: psId = "PR_SENDER_EMAIL_ADDRESS"; break;
  case PR_CURRENT_VERSION >> 16: psId = "PR_CURRENT_VERSION"; break;
  case PR_DELETE_AFTER_SUBMIT >> 16: psId = "PR_DELETE_AFTER_SUBMIT"; break;
  case PR_DISPLAY_BCC >> 16: psId = "PR_DISPLAY_BCC"; break;
  case PR_DISPLAY_CC >> 16: psId = "PR_DISPLAY_CC"; break;
  case PR_DISPLAY_TO >> 16: psId = "PR_DISPLAY_TO"; break;
  case PR_PARENT_DISPLAY >> 16: psId = "PR_PARENT_DISPLAY"; break;
  case PR_MESSAGE_DELIVERY_TIME >> 16: psId = "PR_MESSAGE_DELIVERY_TIME"; break;
  case PR_MESSAGE_FLAGS >> 16: psId = "PR_MESSAGE_FLAGS"; break;
  case PR_MESSAGE_SIZE >> 16: psId = "PR_MESSAGE_SIZE"; break;
  case PR_PARENT_ENTRYID >> 16: psId = "PR_PARENT_ENTRYID"; break;
  case PR_SENTMAIL_ENTRYID >> 16: psId = "PR_SENTMAIL_ENTRYID"; break;
  case PR_CORRELATE >> 16: psId = "PR_CORRELATE"; break;
  case PR_CORRELATE_MTSID >> 16: psId = "PR_CORRELATE_MTSID"; break;
  case PR_DISCRETE_VALUES >> 16: psId = "PR_DISCRETE_VALUES"; break;
  case PR_RESPONSIBILITY >> 16: psId = "PR_RESPONSIBILITY"; break;
  case PR_SPOOLER_STATUS >> 16: psId = "PR_SPOOLER_STATUS"; break;
  case PR_TRANSPORT_STATUS >> 16: psId = "PR_TRANSPORT_STATUS"; break;
  case PR_MESSAGE_RECIPIENTS >> 16: psId = "PR_MESSAGE_RECIPIENTS"; break;
  case PR_MESSAGE_ATTACHMENTS >> 16: psId = "PR_MESSAGE_ATTACHMENTS"; break;
  case PR_SUBMIT_FLAGS >> 16: psId = "PR_SUBMIT_FLAGS"; break;
  case PR_RECIPIENT_STATUS >> 16: psId = "PR_RECIPIENT_STATUS"; break;
  case PR_TRANSPORT_KEY >> 16: psId = "PR_TRANSPORT_KEY"; break;
  case PR_MSG_STATUS >> 16: psId = "PR_MSG_STATUS"; break;
  case PR_MESSAGE_DOWNLOAD_TIME >> 16: psId = "PR_MESSAGE_DOWNLOAD_TIME"; break;
  case PR_CREATION_VERSION >> 16: psId = "PR_CREATION_VERSION"; break;
  case PR_MODIFY_VERSION >> 16: psId = "PR_MODIFY_VERSION"; break;
  case PR_HASATTACH >> 16: psId = "PR_HASATTACH"; break;
  case PR_BODY_CRC >> 16: psId = "PR_BODY_CRC"; break;
  case PR_NORMALIZED_SUBJECT >> 16: psId = "PR_NORMALIZED_SUBJECT"; break;
  case PR_RTF_IN_SYNC >> 16: psId = "PR_RTF_IN_SYNC"; break;
  case PR_ATTACH_SIZE >> 16: psId = "PR_ATTACH_SIZE"; break;
  case PR_ATTACH_NUM >> 16: psId = "PR_ATTACH_NUM"; break;
  case PR_PREPROCESS >> 16: psId = "PR_PREPROCESS"; break;
  case PR_ORIGINATING_MTA_CERTIFICATE >> 16: psId = "PR_ORIGINATING_MTA_CERTIFICATE"; break;
  case PR_PROOF_OF_SUBMISSION >> 16: psId = "PR_PROOF_OF_SUBMISSION"; break;
  case PR_ENTRYID >> 16: psId = "PR_ENTRYID"; break;
  case PR_OBJECT_TYPE >> 16: psId = "PR_OBJECT_TYPE"; break;
  case PR_ICON >> 16: psId = "PR_ICON"; break;
  case PR_MINI_ICON >> 16: psId = "PR_MINI_ICON"; break;
  case PR_STORE_ENTRYID >> 16: psId = "PR_STORE_ENTRYID"; break;
  case PR_STORE_RECORD_KEY >> 16: psId = "PR_STORE_RECORD_KEY"; break;
  case PR_RECORD_KEY >> 16: psId = "PR_RECORD_KEY"; break;
  case PR_MAPPING_SIGNATURE >> 16: psId = "PR_MAPPING_SIGNATURE"; break;
  case PR_ACCESS_LEVEL >> 16: psId = "PR_ACCESS_LEVEL"; break;
  case PR_INSTANCE_KEY >> 16: psId = "PR_INSTANCE_KEY"; break;
  case PR_ROW_TYPE >> 16: psId = "PR_ROW_TYPE"; break;
  case PR_ACCESS >> 16: psId = "PR_ACCESS"; break;
  case PR_ROWID >> 16: psId = "PR_ROWID"; break;
  case PR_DISPLAY_NAME >> 16: psId = "PR_DISPLAY_NAME"; break;
  case PR_ADDRTYPE >> 16: psId = "PR_ADDRTYPE"; break;
  case PR_EMAIL_ADDRESS >> 16: psId = "PR_EMAIL_ADDRESS"; break;
  case PR_COMMENT >> 16: psId = "PR_COMMENT"; break;
  case PR_DEPTH >> 16: psId = "PR_DEPTH"; break;
  case PR_PROVIDER_DISPLAY >> 16: psId = "PR_PROVIDER_DISPLAY"; break;
  case PR_CREATION_TIME >> 16: psId = "PR_CREATION_TIME"; break;
  case PR_LAST_MODIFICATION_TIME >> 16: psId = "PR_LAST_MODIFICATION_TIME"; break;
  case PR_RESOURCE_FLAGS >> 16: psId = "PR_RESOURCE_FLAGS"; break;
  case PR_PROVIDER_DLL_NAME >> 16: psId = "PR_PROVIDER_DLL_NAME"; break;
  case PR_SEARCH_KEY >> 16: psId = "PR_SEARCH_KEY"; break;
  case PR_PROVIDER_UID >> 16: psId = "PR_PROVIDER_UID"; break;
  case PR_PROVIDER_ORDINAL >> 16: psId = "PR_PROVIDER_ORDINAL"; break;
  case PR_FORM_VERSION >> 16: psId = "PR_FORM_VERSION"; break;
  case PR_FORM_CLSID >> 16: psId = "PR_FORM_CLSID"; break;
  case PR_FORM_CONTACT_NAME >> 16: psId = "PR_FORM_CONTACT_NAME"; break;
  case PR_FORM_CATEGORY >> 16: psId = "PR_FORM_CATEGORY"; break;
  case PR_FORM_CATEGORY_SUB >> 16: psId = "PR_FORM_CATEGORY_SUB"; break;
  case PR_FORM_HOST_MAP >> 16: psId = "PR_FORM_HOST_MAP"; break;
  case PR_FORM_HIDDEN >> 16: psId = "PR_FORM_HIDDEN"; break;
  case PR_FORM_DESIGNER_NAME >> 16: psId = "PR_FORM_DESIGNER_NAME"; break;
  case PR_FORM_DESIGNER_GUID >> 16: psId = "PR_FORM_DESIGNER_GUID"; break;
  case PR_FORM_MESSAGE_BEHAVIOR >> 16: psId = "PR_FORM_MESSAGE_BEHAVIOR"; break;
  case PR_DEFAULT_STORE >> 16: psId = "PR_DEFAULT_STORE"; break;
  case PR_STORE_SUPPORT_MASK >> 16: psId = "PR_STORE_SUPPORT_MASK"; break;
  case PR_STORE_STATE >> 16: psId = "PR_STORE_STATE"; break;
  case PR_IPM_SUBTREE_SEARCH_KEY >> 16: psId = "PR_IPM_SUBTREE_SEARCH_KEY"; break;
  case PR_IPM_OUTBOX_SEARCH_KEY >> 16: psId = "PR_IPM_OUTBOX_SEARCH_KEY"; break;
  case PR_IPM_WASTEBASKET_SEARCH_KEY >> 16: psId = "PR_IPM_WASTEBASKET_SEARCH_KEY"; break;
  case PR_IPM_SENTMAIL_SEARCH_KEY >> 16: psId = "PR_IPM_SENTMAIL_SEARCH_KEY"; break;
  case PR_MDB_PROVIDER >> 16: psId = "PR_MDB_PROVIDER"; break;
  case PR_RECEIVE_FOLDER_SETTINGS >> 16: psId = "PR_RECEIVE_FOLDER_SETTINGS"; break;
  case PR_VALID_FOLDER_MASK >> 16: psId = "PR_VALID_FOLDER_MASK"; break;
  case PR_IPM_SUBTREE_ENTRYID >> 16: psId = "PR_IPM_SUBTREE_ENTRYID"; break;
  case PR_IPM_OUTBOX_ENTRYID >> 16: psId = "PR_IPM_OUTBOX_ENTRYID"; break;
  case PR_IPM_WASTEBASKET_ENTRYID >> 16: psId = "PR_IPM_WASTEBASKET_ENTRYID"; break;
  case PR_IPM_SENTMAIL_ENTRYID >> 16: psId = "PR_IPM_SENTMAIL_ENTRYID"; break;
  case PR_VIEWS_ENTRYID >> 16: psId = "PR_VIEWS_ENTRYID"; break;
  case PR_COMMON_VIEWS_ENTRYID >> 16: psId = "PR_COMMON_VIEWS_ENTRYID"; break;
  case PR_FINDER_ENTRYID >> 16: psId = "PR_FINDER_ENTRYID"; break;
  case PR_CONTAINER_FLAGS >> 16: psId = "PR_CONTAINER_FLAGS"; break;
  case PR_FOLDER_TYPE >> 16: psId = "PR_FOLDER_TYPE"; break;
  case PR_CONTENT_COUNT >> 16: psId = "PR_CONTENT_COUNT"; break;
  case PR_CONTENT_UNREAD >> 16: psId = "PR_CONTENT_UNREAD"; break;
  case PR_CREATE_TEMPLATES >> 16: psId = "PR_CREATE_TEMPLATES"; break;
  case PR_DETAILS_TABLE >> 16: psId = "PR_DETAILS_TABLE"; break;
  case PR_SEARCH >> 16: psId = "PR_SEARCH"; break;
  case PR_SELECTABLE >> 16: psId = "PR_SELECTABLE"; break;
  case PR_SUBFOLDERS >> 16: psId = "PR_SUBFOLDERS"; break;
  case PR_STATUS >> 16: psId = "PR_STATUS"; break;
  case PR_ANR >> 16: psId = "PR_ANR"; break;
  case PR_CONTENTS_SORT_ORDER >> 16: psId = "PR_CONTENTS_SORT_ORDER"; break;
  case PR_CONTAINER_HIERARCHY >> 16: psId = "PR_CONTAINER_HIERARCHY"; break;
  case PR_CONTAINER_CONTENTS >> 16: psId = "PR_CONTAINER_CONTENTS"; break;
  case PR_FOLDER_ASSOCIATED_CONTENTS >> 16: psId = "PR_FOLDER_ASSOCIATED_CONTENTS"; break;
  case PR_DEF_CREATE_DL >> 16: psId = "PR_DEF_CREATE_DL"; break;
  case PR_DEF_CREATE_MAILUSER >> 16: psId = "PR_DEF_CREATE_MAILUSER"; break;
  case PR_CONTAINER_CLASS >> 16: psId = "PR_CONTAINER_CLASS"; break;
  case PR_CONTAINER_MODIFY_VERSION >> 16: psId = "PR_CONTAINER_MODIFY_VERSION"; break;
  case PR_AB_PROVIDER_ID >> 16: psId = "PR_AB_PROVIDER_ID"; break;
  case PR_DEFAULT_VIEW_ENTRYID >> 16: psId = "PR_DEFAULT_VIEW_ENTRYID"; break;
  case PR_ASSOC_CONTENT_COUNT >> 16: psId = "PR_ASSOC_CONTENT_COUNT"; break;
  case PR_ATTACHMENT_X400_PARAMETERS >> 16: psId = "PR_ATTACHMENT_X400_PARAMETERS"; break;
  case PR_ATTACH_DATA_BIN >> 16: psId = "PR_ATTACH_DATA"; break;
  // case PR_ATTACH_DATA_OBJ >> 16: psId = "PR_ATTACH_DATA"; break;
  case PR_ATTACH_ENCODING >> 16: psId = "PR_ATTACH_ENCODING"; break;
  case PR_ATTACH_EXTENSION >> 16: psId = "PR_ATTACH_EXTENSION"; break;
  case PR_ATTACH_FILENAME >> 16: psId = "PR_ATTACH_FILENAME"; break;
  case PR_ATTACH_METHOD >> 16: psId = "PR_ATTACH_METHOD"; break;
  case PR_ATTACH_LONG_FILENAME >> 16: psId = "PR_ATTACH_LONG_FILENAME"; break;
  case PR_ATTACH_PATHNAME >> 16: psId = "PR_ATTACH_PATHNAME"; break;
  case PR_ATTACH_RENDERING >> 16: psId = "PR_ATTACH_RENDERING"; break;
  case PR_ATTACH_TAG >> 16: psId = "PR_ATTACH_TAG"; break;
  case PR_RENDERING_POSITION >> 16: psId = "PR_RENDERING_POSITION"; break;
  case PR_ATTACH_TRANSPORT_NAME >> 16: psId = "PR_ATTACH_TRANSPORT_NAME"; break;
  case PR_ATTACH_LONG_PATHNAME >> 16: psId = "PR_ATTACH_LONG_PATHNAME"; break;
  case PR_ATTACH_MIME_TAG >> 16: psId = "PR_ATTACH_MIME_TAG"; break;
  case PR_ATTACH_ADDITIONAL_INFO >> 16: psId = "PR_ATTACH_ADDITIONAL_INFO"; break;
  case PR_DISPLAY_TYPE >> 16: psId = "PR_DISPLAY_TYPE"; break;
  case PR_TEMPLATEID >> 16: psId = "PR_TEMPLATEID"; break;
  case PR_PRIMARY_CAPABILITY >> 16: psId = "PR_PRIMARY_CAPABILITY"; break;
  case PR_7BIT_DISPLAY_NAME >> 16: psId = "PR_7BIT_DISPLAY_NAME"; break;
  case PR_ACCOUNT >> 16: psId = "PR_ACCOUNT"; break;
  case PR_ALTERNATE_RECIPIENT >> 16: psId = "PR_ALTERNATE_RECIPIENT"; break;
  case PR_CALLBACK_TELEPHONE_NUMBER >> 16: psId = "PR_CALLBACK_TELEPHONE_NUMBER"; break;
  case PR_CONVERSION_PROHIBITED >> 16: psId = "PR_CONVERSION_PROHIBITED"; break;
  case PR_DISCLOSE_RECIPIENTS >> 16: psId = "PR_DISCLOSE_RECIPIENTS"; break;
  case PR_GENERATION >> 16: psId = "PR_GENERATION"; break;
  case PR_GIVEN_NAME >> 16: psId = "PR_GIVEN_NAME"; break;
  case PR_GOVERNMENT_ID_NUMBER >> 16: psId = "PR_GOVERNMENT_ID_NUMBER"; break;
  case PR_BUSINESS_TELEPHONE_NUMBER >> 16: psId = "PR_BUSINESS_TELEPHONE_NUMBER"; break;
  case PR_HOME_TELEPHONE_NUMBER >> 16: psId = "PR_HOME_TELEPHONE_NUMBER"; break;
  case PR_INITIALS >> 16: psId = "PR_INITIALS"; break;
  case PR_KEYWORD >> 16: psId = "PR_KEYWORD"; break;
  case PR_LANGUAGE >> 16: psId = "PR_LANGUAGE"; break;
  case PR_LOCATION >> 16: psId = "PR_LOCATION"; break;
  case PR_MAIL_PERMISSION >> 16: psId = "PR_MAIL_PERMISSION"; break;
  case PR_MHS_COMMON_NAME >> 16: psId = "PR_MHS_COMMON_NAME"; break;
  case PR_ORGANIZATIONAL_ID_NUMBER >> 16: psId = "PR_ORGANIZATIONAL_ID_NUMBER"; break;
  case PR_SURNAME >> 16: psId = "PR_SURNAME"; break;
  case PR_ORIGINAL_ENTRYID >> 16: psId = "PR_ORIGINAL_ENTRYID"; break;
  case PR_ORIGINAL_DISPLAY_NAME >> 16: psId = "PR_ORIGINAL_DISPLAY_NAME"; break;
  case PR_ORIGINAL_SEARCH_KEY >> 16: psId = "PR_ORIGINAL_SEARCH_KEY"; break;
  case PR_POSTAL_ADDRESS >> 16: psId = "PR_POSTAL_ADDRESS"; break;
  case PR_COMPANY_NAME >> 16: psId = "PR_COMPANY_NAME"; break;
  case PR_TITLE >> 16: psId = "PR_TITLE"; break;
  case PR_DEPARTMENT_NAME >> 16: psId = "PR_DEPARTMENT_NAME"; break;
  case PR_OFFICE_LOCATION >> 16: psId = "PR_OFFICE_LOCATION"; break;
  case PR_PRIMARY_TELEPHONE_NUMBER >> 16: psId = "PR_PRIMARY_TELEPHONE_NUMBER"; break;
  case PR_BUSINESS2_TELEPHONE_NUMBER >> 16: psId = "PR_BUSINESS2_TELEPHONE_NUMBER"; break;
  case PR_MOBILE_TELEPHONE_NUMBER >> 16: psId = "PR_MOBILE_TELEPHONE_NUMBER"; break;
  case PR_RADIO_TELEPHONE_NUMBER >> 16: psId = "PR_RADIO_TELEPHONE_NUMBER"; break;
  case PR_CAR_TELEPHONE_NUMBER >> 16: psId = "PR_CAR_TELEPHONE_NUMBER"; break;
  case PR_OTHER_TELEPHONE_NUMBER >> 16: psId = "PR_OTHER_TELEPHONE_NUMBER"; break;
  case PR_TRANSMITABLE_DISPLAY_NAME >> 16: psId = "PR_TRANSMITABLE_DISPLAY_NAME"; break;
  case PR_PAGER_TELEPHONE_NUMBER >> 16: psId = "PR_PAGER_TELEPHONE_NUMBER"; break;
  case PR_USER_CERTIFICATE >> 16: psId = "PR_USER_CERTIFICATE"; break;
  case PR_PRIMARY_FAX_NUMBER >> 16: psId = "PR_PRIMARY_FAX_NUMBER"; break;
  case PR_BUSINESS_FAX_NUMBER >> 16: psId = "PR_BUSINESS_FAX_NUMBER"; break;
  case PR_HOME_FAX_NUMBER >> 16: psId = "PR_HOME_FAX_NUMBER"; break;
  case PR_COUNTRY >> 16: psId = "PR_COUNTRY"; break;
  case PR_LOCALITY >> 16: psId = "PR_LOCALITY"; break;
  case PR_STATE_OR_PROVINCE >> 16: psId = "PR_STATE_OR_PROVINCE"; break;
  case PR_STREET_ADDRESS >> 16: psId = "PR_STREET_ADDRESS"; break;
  case PR_POSTAL_CODE >> 16: psId = "PR_POSTAL_CODE"; break;
  case PR_POST_OFFICE_BOX >> 16: psId = "PR_POST_OFFICE_BOX"; break;
  case PR_TELEX_NUMBER >> 16: psId = "PR_TELEX_NUMBER"; break;
  case PR_ISDN_NUMBER >> 16: psId = "PR_ISDN_NUMBER"; break;
  case PR_ASSISTANT_TELEPHONE_NUMBER >> 16: psId = "PR_ASSISTANT_TELEPHONE_NUMBER"; break;
  case PR_HOME2_TELEPHONE_NUMBER >> 16: psId = "PR_HOME2_TELEPHONE_NUMBER"; break;
  case PR_ASSISTANT >> 16: psId = "PR_ASSISTANT"; break;
  case PR_SEND_RICH_INFO >> 16: psId = "PR_SEND_RICH_INFO"; break;
  case PR_WEDDING_ANNIVERSARY >> 16: psId = "PR_WEDDING_ANNIVERSARY"; break;
  case PR_BIRTHDAY >> 16: psId = "PR_BIRTHDAY"; break;
  case PR_HOBBIES >> 16: psId = "PR_HOBBIES"; break;
  case PR_MIDDLE_NAME >> 16: psId = "PR_MIDDLE_NAME"; break;
  case PR_DISPLAY_NAME_PREFIX >> 16: psId = "PR_DISPLAY_NAME_PREFIX"; break;
  case PR_PROFESSION >> 16: psId = "PR_PROFESSION"; break;
  case PR_PREFERRED_BY_NAME >> 16: psId = "PR_PREFERRED_BY_NAME"; break;
  case PR_SPOUSE_NAME >> 16: psId = "PR_SPOUSE_NAME"; break;
  case PR_COMPUTER_NETWORK_NAME >> 16: psId = "PR_COMPUTER_NETWORK_NAME"; break;
  case PR_CUSTOMER_ID >> 16: psId = "PR_CUSTOMER_ID"; break;
  case PR_TTYTDD_PHONE_NUMBER >> 16: psId = "PR_TTYTDD_PHONE_NUMBER"; break;
  case PR_FTP_SITE >> 16: psId = "PR_FTP_SITE"; break;
  case PR_GENDER >> 16: psId = "PR_GENDER"; break;
  case PR_MANAGER_NAME >> 16: psId = "PR_MANAGER_NAME"; break;
  case PR_NICKNAME >> 16: psId = "PR_NICKNAME"; break;
  case PR_PERSONAL_HOME_PAGE >> 16: psId = "PR_PERSONAL_HOME_PAGE"; break;
  case PR_BUSINESS_HOME_PAGE >> 16: psId = "PR_BUSINESS_HOME_PAGE"; break;
  case PR_CONTACT_VERSION >> 16: psId = "PR_CONTACT_VERSION"; break;
  case PR_CONTACT_ENTRYIDS >> 16: psId = "PR_CONTACT_ENTRYIDS"; break;
  case PR_CONTACT_ADDRTYPES >> 16: psId = "PR_CONTACT_ADDRTYPES"; break;
  case PR_CONTACT_DEFAULT_ADDRESS_INDEX >> 16: psId = "PR_CONTACT_DEFAULT_ADDRESS_INDEX"; break;
  case PR_CONTACT_EMAIL_ADDRESSES >> 16: psId = "PR_CONTACT_EMAIL_ADDRESSES"; break;
  case PR_COMPANY_MAIN_PHONE_NUMBER >> 16: psId = "PR_COMPANY_MAIN_PHONE_NUMBER"; break;
  case PR_CHILDRENS_NAMES >> 16: psId = "PR_CHILDRENS_NAMES"; break;
  case PR_HOME_ADDRESS_CITY >> 16: psId = "PR_HOME_ADDRESS_CITY"; break;
  case PR_HOME_ADDRESS_COUNTRY >> 16: psId = "PR_HOME_ADDRESS_COUNTRY"; break;
  case PR_HOME_ADDRESS_POSTAL_CODE >> 16: psId = "PR_HOME_ADDRESS_POSTAL_CODE"; break;
  case PR_HOME_ADDRESS_STATE_OR_PROVINCE >> 16: psId = "PR_HOME_ADDRESS_STATE_OR_PROVINCE"; break;
  case PR_HOME_ADDRESS_STREET >> 16: psId = "PR_HOME_ADDRESS_STREET"; break;
  case PR_HOME_ADDRESS_POST_OFFICE_BOX >> 16: psId = "PR_HOME_ADDRESS_POST_OFFICE_BOX"; break;
  case PR_OTHER_ADDRESS_CITY >> 16: psId = "PR_OTHER_ADDRESS_CITY"; break;
  case PR_OTHER_ADDRESS_COUNTRY >> 16: psId = "PR_OTHER_ADDRESS_COUNTRY"; break;
  case PR_OTHER_ADDRESS_POSTAL_CODE >> 16: psId = "PR_OTHER_ADDRESS_POSTAL_CODE"; break;
  case PR_OTHER_ADDRESS_STATE_OR_PROVINCE >> 16: psId = "PR_OTHER_ADDRESS_STATE_OR_PROVINCE"; break;
  case PR_OTHER_ADDRESS_STREET >> 16: psId = "PR_OTHER_ADDRESS_STREET"; break;
  case PR_OTHER_ADDRESS_POST_OFFICE_BOX >> 16: psId = "PR_OTHER_ADDRESS_POST_OFFICE_BOX"; break;
  case PR_STORE_PROVIDERS >> 16: psId = "PR_STORE_PROVIDERS"; break;
  case PR_AB_PROVIDERS >> 16: psId = "PR_AB_PROVIDERS"; break;
  case PR_TRANSPORT_PROVIDERS >> 16: psId = "PR_TRANSPORT_PROVIDERS"; break;
  case PR_DEFAULT_PROFILE >> 16: psId = "PR_DEFAULT_PROFILE"; break;
  case PR_AB_SEARCH_PATH >> 16: psId = "PR_AB_SEARCH_PATH"; break;
  case PR_AB_DEFAULT_DIR >> 16: psId = "PR_AB_DEFAULT_DIR"; break;
  case PR_AB_DEFAULT_PAB >> 16: psId = "PR_AB_DEFAULT_PAB"; break;
  case PR_FILTERING_HOOKS >> 16: psId = "PR_FILTERING_HOOKS"; break;
  case PR_SERVICE_NAME >> 16: psId = "PR_SERVICE_NAME"; break;
  case PR_SERVICE_DLL_NAME >> 16: psId = "PR_SERVICE_DLL_NAME"; break;
  case PR_SERVICE_ENTRY_NAME >> 16: psId = "PR_SERVICE_ENTRY_NAME"; break;
  case PR_SERVICE_UID >> 16: psId = "PR_SERVICE_UID"; break;
  case PR_SERVICE_EXTRA_UIDS >> 16: psId = "PR_SERVICE_EXTRA_UIDS"; break;
  case PR_SERVICES >> 16: psId = "PR_SERVICES"; break;
  case PR_SERVICE_SUPPORT_FILES >> 16: psId = "PR_SERVICE_SUPPORT_FILES"; break;
  case PR_SERVICE_DELETE_FILES >> 16: psId = "PR_SERVICE_DELETE_FILES"; break;
  case PR_AB_SEARCH_PATH_UPDATE >> 16: psId = "PR_AB_SEARCH_PATH_UPDATE"; break;
  case PR_PROFILE_NAME >> 16: psId = "PR_PROFILE_NAME"; break;
  case PR_IDENTITY_DISPLAY >> 16: psId = "PR_IDENTITY_DISPLAY"; break;
  case PR_IDENTITY_ENTRYID >> 16: psId = "PR_IDENTITY_ENTRYID"; break;
  case PR_RESOURCE_METHODS >> 16: psId = "PR_RESOURCE_METHODS"; break;
  case PR_RESOURCE_TYPE >> 16: psId = "PR_RESOURCE_TYPE"; break;
  case PR_STATUS_CODE >> 16: psId = "PR_STATUS_CODE"; break;
  case PR_IDENTITY_SEARCH_KEY >> 16: psId = "PR_IDENTITY_SEARCH_KEY"; break;
  case PR_OWN_STORE_ENTRYID >> 16: psId = "PR_OWN_STORE_ENTRYID"; break;
  case PR_RESOURCE_PATH >> 16: psId = "PR_RESOURCE_PATH"; break;
  case PR_STATUS_STRING >> 16: psId = "PR_STATUS_STRING"; break;
  case PR_X400_DEFERRED_DELIVERY_CANCEL >> 16: psId = "PR_X400_DEFERRED_DELIVERY_CANCEL"; break;
  case PR_HEADER_FOLDER_ENTRYID >> 16: psId = "PR_HEADER_FOLDER_ENTRYID"; break;
  case PR_REMOTE_PROGRESS >> 16: psId = "PR_REMOTE_PROGRESS"; break;
  case PR_REMOTE_PROGRESS_TEXT >> 16: psId = "PR_REMOTE_PROGRESS_TEXT"; break;
  case PR_REMOTE_VALIDATE_OK >> 16: psId = "PR_REMOTE_VALIDATE_OK"; break;
  case PR_CONTROL_FLAGS >> 16: psId = "PR_CONTROL_FLAGS"; break;
  case PR_CONTROL_STRUCTURE >> 16: psId = "PR_CONTROL_STRUCTURE"; break;
  case PR_CONTROL_TYPE >> 16: psId = "PR_CONTROL_TYPE"; break;
  case PR_DELTAX >> 16: psId = "PR_DELTAX"; break;
  case PR_DELTAY >> 16: psId = "PR_DELTAY"; break;
  case PR_XPOS >> 16: psId = "PR_XPOS"; break;
  case PR_YPOS >> 16: psId = "PR_YPOS"; break;
  case PR_CONTROL_ID >> 16: psId = "PR_CONTROL_ID"; break;
  case PR_INITIAL_DETAILS_PANE >> 16: psId = "PR_INITIAL_DETAILS_PANE"; break;
  default: psId ="PR_UNKNOWN"; break;
  }
  return psId;
} /* id */

/*--------------------------------------------------------------------*/
static char acValue[256];
char* PropValue::value()
{
  sprintf(acValue,"0x%08x",&m_pSPropValue->Value.ul);
  char* psValue = acValue;
  USHORT usPropType = (USHORT)(this->tag() & 0x0000FFFF);
  switch (usPropType)
  {
  case PT_UNSPECIFIED: break;
  case PT_NULL: break;
  case PT_I2: sprintf(acValue,"%d",(int)m_pSPropValue->Value.i); break;
  case PT_I4: sprintf(acValue,"%ld",(LONG)m_pSPropValue->Value.l); break;
  case PT_I8: sprintf(acValue,"0x%16x",m_pSPropValue->Value.li); break;
  case PT_R4: sprintf(acValue,"%f",(double)m_pSPropValue->Value.flt); break;
  case PT_R8: sprintf(acValue,"%f", m_pSPropValue->Value.dbl); break;
  case PT_CURRENCY: break;
  case PT_APPTIME: break;
  case PT_ERROR: sprintf(acValue,"0x%08x",m_pSPropValue->Value.ul); break;
  case PT_BOOLEAN: sprintf(acValue,m_pSPropValue->Value.i == 0? "false": "true"); break;
  case PT_OBJECT: break;
  case PT_STRING8: psValue = m_pSPropValue->Value.lpszA; break;
  case PT_UNICODE: break;
  case PT_SYSTIME: break;
  case PT_CLSID: psValue = m_pSPropValue->Value.lpszA; break;
  case PT_BINARY: break;
  default: break;
  }
  return psValue;
} /* value */

BOOL PropValue::getBooleanValue()
{
  return m_pSPropValue->Value.i == 0? FALSE: TRUE;
} /* getBooleanValue */

/*--------------------------------------------------------------------*/
ULONG PropValue::getLongValue()
{
  return m_pSPropValue->Value.ul;
} /* getLongValue */

/*--------------------------------------------------------------------*/
char* PropValue::getString8Value()
{
  return _strdup(m_pSPropValue->Value.lpszA);
} /* getString8Value */

/*--------------------------------------------------------------------*/
LPBYTE PropValue::getBinaryPointer()
{
  return m_pSPropValue->Value.bin.lpb;
} /* getBinaryPointer */

/*--------------------------------------------------------------------*/
ULONG PropValue::getBinaryLength()
{
  return m_pSPropValue->Value.bin.cb;
} /* getBinaryLength */

/*--------------------------------------------------------------------*/
FILETIME* PropValue::getSysTimeValue()
{
  return &m_pSPropValue->Value.ft;
} /* getSysTimeValue */

/*--------------------------------------------------------------------*/
HRESULT PropValue::list()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  ULONG ulPropTag = this->tag();
  printf("%s (0x%04x) of type: %s (0x%04x): %s\n",
    this->id(),(USHORT)(ulPropTag >> 16),
    this->type(),(USHORT)(ulPropTag & 0x0000FFFF),
    this->value());
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* list */

/*==End of File=======================================================*/
