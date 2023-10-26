import React, { useState } from "react";
import CreateWidgetForCardUI from "./CreateFormFunc";
import PropTypes from "prop-types";
import { setOfficeKeyValue, officeKeys } from "../../config/utility";
import { useSelector } from "react-redux";
/* globals Office, fetch */
function DynamicForm(props) {
  const { selectedFormDetails, domain, orgId, setRequestId } = props;
  const [values, setValues] = useState({});
  const [attachmentData, setAttachmentData] = useState({});
  const authData = useSelector((state) => state.auth);

  const updateTextField = (event) => {
    const { name, value } = event.target;
    setValues({ ...values, [name]: value });
  };
  setOfficeKeyValue(officeKeys.selectedFormDetails, JSON.stringify(selectedFormDetails));

  const readData = (attachmentId) => {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, { asyncContext: null }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const attachmentContent = result.value;
          setAttachmentData(attachmentContent);
          resolve(attachmentContent);
        } else {
          // Handle errors
          reject(result.error);
        }
      });
    });
  };

  const handleSubmit = async () => {
    let attachmentContent;

    if (attachmentData?.id) {
      attachmentContent = await readData(attachmentData.id);
    }
    const payload = {
      answers: values,
      formId: selectedFormDetails._id,
      attachments: attachmentData?.id ? [{ ...attachmentData, ...attachmentContent }] : [],
      messageId: Office.context.mailbox.item.internetMessageId,
      subject: Office.context.mailbox.item.subject,
    };
    const res = await fetch(`${domain}/api/outlook/create-request/${orgId}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${authData.accessToken}`,
      },
      body: JSON.stringify(payload),
    });
    if (res.status == 200) {
      const data = await res.json();
      const conversation = Office.context.mailbox.initialData.conversationId;
      setOfficeKeyValue(conversation, data._id);
      setRequestId(data._id);
    }
  };

  const getAttachmentData = (attachment) => {
    setAttachmentData(attachment);
  };

  return (
    <>
      <p
        style={{
          fontWeight: "bold",
          fontFamily: "sans-serif",
          fontSize: "20px",
          padding: "10px 10px",
          marginLeft: "10px",
          border: "1px solid #ccc",
        }}
      >
        {selectedFormDetails.matter} Form
      </p>{" "}
      {selectedFormDetails.fields?.length &&
        selectedFormDetails.fields.map((item) => {
          return (
            <>
              <div style={{ marginTop: "20px", marginLeft: "10px" }}>
                <p
                  style={{
                    fontSize: "15px",
                    width: "90%",
                    fontFamily: "sans-serif",
                    fontWeight: "bold",
                  }}
                >
                  {item.title}{" "}
                </p>{" "}
                <CreateWidgetForCardUI
                  field={item}
                  onChange={updateTextField}
                  values={values}
                  attachments={Office.context.mailbox.item.attachments}
                  getAttachmentData={getAttachmentData}
                />{" "}
              </div>
            </>
          );
        })}
      <button
        style={{
          marginLeft: "10px",
          width: "50%",
          marginTop: "10px",
          backgroundColor: "white",
          border: "1px solid #ccc",
          fontSize: "16px",
          borderRadius: "2px",
          padding: "10px 0px",
          cursor: "pointer",
        }}
        onClick={handleSubmit}
      >
        Submit
      </button>
    </>
  );
}

DynamicForm.propTypes = {
  selectedFormDetails: PropTypes.object.isRequired,
  domain: PropTypes.string.isRequired,
  orgId: PropTypes.string.isRequired,
  setRequestId: PropTypes.func.isRequired,
};
export default DynamicForm;
