/* globals Office, fetch */
import React, { useState } from "react";
import CreateWidgetForCardUI from "./CreateFormFunc";
import PropTypes from "prop-types";
import { setOfficeKeyValue, officeKeys } from "../../config/utility";
import { useDispatch, useSelector } from "react-redux";
import { setAlertMessage, setLoading } from "../../app/loaderSlice";
import { Text, Button, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    marginTop: "20px",
    marginLeft: "10px",
    display: "flex",
    flexDirection: "column",
  },
  text: {
    fontSize: "16px",
    width: "90%",
    fontFamily: "sans-serif",
    marginBottom: "5px",
  },
  button: {
    marginTop: "10px",
    marginLeft: "10px",
    width: "50%",
    fontSize: "16px",
  },
  title: {
    fontWeight: "bold",
    fontFamily: "sans-serif",
    fontSize: "20px",
    marginLeft: "10px",
  },
});

function DynamicForm(props) {
  const { selectedFormDetails, domain, orgId, setRequestId } = props;
  const [values, setValues] = useState({});
  const [attachmentData, setAttachmentData] = useState({});
  const authData = useSelector((state) => state.auth);
  const dispatchToRedux = useDispatch();
  const classes = useStyles();

  const updateTextField = (event) => {
    const { name, value } = event.target;
    if (event?.target.type === "checkbox") {
      if (!values[name]) {
        setValues({ ...values, [name]: [value] });
      } else if (!values[name].includes(value)) {
        setValues({ ...values, [name]: [...values[name], value] });
      }
    } else {
      setValues({ ...values, [name]: value });
    }
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
          reject(result.error);
        }
      });
    });
  };

  const handleSubmit = async () => {
    let attachmentContent;
    dispatchToRedux(setLoading(true));

    try {
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
      } else {
        const errorData = await res.json();
        dispatchToRedux(setAlertMessage({ message: errorData.message, intent: "error" }));
      }
    } catch (error) {
      dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
    }
    dispatchToRedux(setLoading(false));
  };

  const getAttachmentData = (attachment) => {
    setAttachmentData(attachment);
  };

  return (
    <>
      <Text className={classes.title}>{selectedFormDetails.matter} Form</Text>{" "}
      {selectedFormDetails.fields?.length &&
        selectedFormDetails.fields.map((item) => {
          return (
            <CreateWidgetForCardUI
              key={item.title}
              field={item}
              onChange={updateTextField}
              values={values}
              attachments={Office.context.mailbox.item.attachments}
              getAttachmentData={getAttachmentData}
            />
          );
        })}
      <Button appearance="primary" className={classes.button} onClick={handleSubmit}>
        Submit
      </Button>
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
