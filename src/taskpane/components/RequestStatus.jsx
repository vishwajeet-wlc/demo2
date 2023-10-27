import React, { useState } from "react";
import { useEffect } from "react";
import PropTypes from "prop-types";
import { getOfficeKeyValue, officeKeys } from "../../config/utility";
import { useDispatch, useSelector } from "react-redux";
import { Text, Input, Select, Button, makeStyles } from "@fluentui/react-components";
import { setAlertMessage, setLoading } from "../../app/loaderSlice";
/* global Office, fetch */

const useStyles = makeStyles({
  text: {
    fontWeight: "bold",
    fontFamily: "sans-serif",
    fontSize: "18px",
    marginLeft: "10px",
  },
  input: {
    width: "80%",
    fontSize: "14px",
    backgroundColor: "#eee",
  },
  select: {
    width: "80%",
    height: "30px",
    marginTop: "20px",
  },
  button: {
    marginTop: "10px",
    backgroundColor: "white",
    fontSize: "16px",
    cursor: "pointer",
    marginLeft: "10px",
  },
  modal: {
    marginLeft: "10px",
    display: "flex",
    flexDirection: "column",
    rowGap: "10px",
  },
});

function RequestStatus({ reqId, clientToken, clientDomain }) {
  const [requestDetail, setRequestDetail] = useState(null);
  const [status, setStatus] = useState("");
  const [isOpen, setIsOpen] = useState(false);
  const authData = useSelector((state) => state.auth);
  const REQUEST_STATUS = ["submitted", "review", "waitingOn", "approved", "completed", "rejected", "cancelled"];
  const classes = useStyles();
  const dispatchToRedux = useDispatch();

  useEffect(() => {
    async function fetchStatus() {
      dispatchToRedux(setLoading(true));
      try {
        const res = await fetch(`${clientDomain}/api/outlook/get-request/${reqId}`, {
          headers: {
            Authorization: `Bearer ${authData.accessToken}`,
          },
        });
        const reqData = await res.json();
        setRequestDetail(reqData);
        setStatus(reqData.status);
      } catch (error) {
        dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
      }
      dispatchToRedux(setLoading(false));
    }

    if (reqId && clientDomain && !requestDetail) {
      fetchStatus();
    }
  }, [requestDetail, reqId, clientDomain]);
  const questions = JSON.parse(getOfficeKeyValue(officeKeys.selectedFormDetails));

  const handleChange = async (e) => {
    dispatchToRedux(setLoading(true));
    const clientToken = getOfficeKeyValue(officeKeys.clientToken);
    const payload = {
      organizationId: clientToken.split(".").pop(),
      updates: {
        status: e.target.value,
      },
    };
    try {
      const response = await fetch(`${clientDomain}/api/outlook/update-request/${reqId}`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${authData.accessToken}`,
        },
        body: JSON.stringify(payload),
      });
      const updatedRequest = await response.json();
      if (updatedRequest.status !== status) {
        setStatus(updatedRequest.status);
        dispatchToRedux(setAlertMessage({ message: "Updated successfully", intent: "success" }));
      }
    } catch (error) {
      dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
    }
    dispatchToRedux(setLoading(false));
  };

  async function getAttachments(attachmentsData) {
    const attachments = await Promise.all(
      attachmentsData.map((attachment) => {
        // Return a promise for each attachment
        return new Promise((resolve, reject) => {
          Office.context.mailbox.item.getAttachmentContentAsync(
            attachment.id,
            { asyncContext: null },
            function (result) {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Resolve the promise with the result value
                resolve({ ...result.value, ...attachment });
              } else {
                // Reject the promise with the result error
                reject(result.error);
              }
            }
          );
        });
      })
    );
    return attachments;
  }

  async function sendToStreamlineAsComment() {
    dispatchToRedux(setLoading(true));
    let htmlBody = await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(asyncResult.value);
          } else {
            reject(asyncResult.error);
          }
        }
      );
    });

    try {
      const attachments = await getAttachments(Office.context.mailbox.item.attachments);
      await fetch(`${clientDomain}/api/outlook/update-request/${reqId}`, {
        method: "PATCH",
        body: JSON.stringify({
          updates: {
            message: `<div>${htmlBody}</div>`,
            messageType: "email",
            emailData: {
              createdViaEmail: true,
              from: Office.context.mailbox.item.from.emailAddress,
              destinations: [
                ...new Set([
                  ...Office.context.mailbox.item.to.map((toData) => toData.emailAddress),
                  ...Office.context.mailbox.item.cc.map((ccEmailData) => ccEmailData.emailAddress),
                ]),
              ],
            },
          },
          attachmentsForMessage: attachments,
          organizationId: clientToken.split(".").pop(),
        }),
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${authData.accessToken}`,
        },
      });
      dispatchToRedux(setAlertMessage({ message: "comment created successfully", intent: "success" }));
    } catch (error) {
      dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
    }
    dispatchToRedux(setLoading(false));
    setIsOpen(false);
  }

  return (
    <>
      {requestDetail && !isOpen && (
        <>
          <Text className={classes.text}>{requestDetail.matter}</Text>
          <div style={{ marginLeft: "10px" }}>
            {questions.fields.map((quest) => {
              return (
                <>
                  {quest.type == "attachment" ? (
                    <div key={quest.title}>
                      <label
                        htmlFor="name"
                        style={{
                          fontSize: "18px",
                          fontFamily: "sans-serif",
                          fontWeigh: "bold",
                          marginTop: "30px",
                        }}
                      ></label>{" "}
                      <br />
                    </div>
                  ) : (
                    <div key={quest.title}>
                      <div style={{ marginTop: "20px" }}>
                        <label
                          htmlFor="name"
                          style={{
                            fontSize: "15px",
                            width: "90%",
                            fontFamily: "sans-serif",
                          }}
                        >
                          {quest.title}{" "}
                        </label>
                      </div>
                      {requestDetail.answers[quest._id] && (
                        <>
                          <Input
                            type="text"
                            disabled
                            className={classes.input}
                            value={requestDetail.answers[quest._id].value}
                            readOnly
                          />{" "}
                          <br />
                        </>
                      )}
                    </div>
                  )}
                </>
              );
            })}
            <label
              htmlFor="name"
              style={{
                fontSize: "18px",
                fontFamily: "sans-serif",
                fontWeigh: "bold",
                marginTop: "20px",
              }}
            >
              Status
            </label>{" "}
            <br />
            <Select className={classes.select} value={status} onChange={handleChange}>
              {REQUEST_STATUS.map((opt) => {
                return (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                );
              })}
            </Select>
          </div>
          <div>
            <Button className={classes.button} onClick={() => setIsOpen(true)}>
              Send To Streamline
            </Button>
          </div>
        </>
      )}

      {isOpen && (
        <div className={classes.modal}>
          <Text> Do you want send this message as comment on streamline ?</Text>
          <div>
            <Button onClick={sendToStreamlineAsComment} appearance="primary">
              Send
            </Button>{" "}
            <Button appearance="secondary" onClick={() => setIsOpen(false)}>
              Cancel
            </Button>
          </div>
        </div>
      )}
    </>
  );
}
RequestStatus.propTypes = {
  reqId: PropTypes.string.isRequired,
  clientToken: PropTypes.string.isRequired,
  clientDomain: PropTypes.string.isRequired,
};
export default RequestStatus;
