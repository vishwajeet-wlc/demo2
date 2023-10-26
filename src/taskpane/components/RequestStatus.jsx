import React, { useState } from "react";
import { useEffect } from "react";
import PropTypes from "prop-types";
import { getOfficeKeyValue, officeKeys } from "../../config/utility";
import { useSelector } from "react-redux";
/* global Office, fetch */
function RequestStatus({ reqId, clientToken, clientDomain }) {
  const [requestDetail, setRequestDetail] = useState(null);
  const [status, setStatus] = useState("");
  const [isOpen, setIsOpen] = useState(false);
  const authData = useSelector((state) => state.auth);
  const REQUEST_STATUS = ["submitted", "review", "waitingOn", "approved", "completed", "rejected", "cancelled"];

  useEffect(() => {
    async function fetchStatus() {
      const res = await fetch(`${clientDomain}/api/outlook/get-request/${reqId}`, {
        headers: {
          Authorization: `Bearer ${authData.accessToken}`,
        },
      });
      const reqData = await res.json();
      setRequestDetail(reqData);
      setStatus(reqData.status);
    }
    if (reqId && clientDomain && !requestDetail) {
      fetchStatus();
    }
  }, [requestDetail, reqId, clientDomain]);
  const questions = JSON.parse(getOfficeKeyValue(officeKeys.selectedFormDetails));

  const handleChange = async (e) => {
    const clientToken = getOfficeKeyValue(officeKeys.clientToken);
    const newStatus = e.target.value;
    setStatus(newStatus);
    const payload = {
      organizationId: clientToken.split(".").pop(),
      updates: {
        status: e.target.value,
        currentHolder: "requestor",
      },
    };
    await fetch(`${clientDomain}/api/outlook/update-request/${reqId}`, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${authData.accessToken}`,
      },
      body: JSON.stringify(payload),
    });
  };

  async function sendToStreamlineAsComment() {
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
        // attachmentsForMessage: Office.context.mailbox.item.attachments,
        organizationId: clientToken.split(".").pop(),
      }),
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${authData.accessToken}`,
      },
    });
    setIsOpen(false);
  }

  return (
    <>
      {requestDetail && !isOpen && (
        <>
          <p
            style={{
              fontWeight: "bold",
              fontFamily: "sans-serif",
              fontSize: "18px",
              padding: "10px 20px",
              border: "1px solid #ccc",
            }}
          >
            {requestDetail.matter}{" "}
          </p>
          <div style={{ marginTop: "10px", marginLeft: "10px" }}>
            {questions.fields.map((quest) => {
              return (
                <>
                  {quest.type == "attachment" ? (
                    <>
                      <label
                        htmlFor="name"
                        style={{ fontSize: "18px", fontFamily: "sans-serif", fontWeigh: "bold", marginTop: "30px" }}
                      ></label>{" "}
                      <br />
                    </>
                  ) : (
                    <div style={{}}>
                      <div style={{ marginTop: "20px" }}>
                        <label
                          htmlFor="name"
                          style={{
                            fontSize: "15px",
                            width: "90%",
                            fontFamily: "sans-serif",
                            fontWeight: "bold",
                            padding: "10px 0px",
                          }}
                        >
                          {quest.title}{" "}
                        </label>
                      </div>
                      {requestDetail.answers[quest._id] && (
                        <>
                          <input
                            type="text"
                            disabled
                            style={{
                              marginRight: "5%",
                              width: "80%",
                              border: "1px solid #ccc",
                              borderRadius: "2px",
                              fontSize: "14px",
                              padding: "10px 10px",
                              backgroundColor: "#eee",
                            }}
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
              style={{ fontSize: "18px", fontFamily: "sans-serif", fontWeigh: "bold", marginTop: "20px" }}
            >
              Status
            </label>{" "}
            <br />
            {/* <input type="text" style={{width:"80%",margin:"auto",height:"30px",marginTop:"20px",fontSize:"18px"}} value={requestDetail.status} readOnly />  */}
            <select
              style={{ width: "90%", margin: "auto", height: "30px", marginTop: "20px" }}
              value={status}
              onChange={handleChange}
            >
              {REQUEST_STATUS.map((opt) => {
                return (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                );
              })}
            </select>
          </div>
          <div>
            <button
              style={{
                marginLeft: "10px",
                marginTop: "10px",
                backgroundColor: "white",
                border: "1px solid #ccc",
                fontSize: "16px",
                borderRadius: "2px",
                padding: "10px 10px",
                cursor: "pointer",
              }}
              onClick={() => setIsOpen(true)}
            >
              Send To Streamline
            </button>
          </div>
        </>
      )}

      {isOpen && (
        <div>
          <p> Do you want send this message as comment on streamline ?</p>
          <button
            style={{
              marginLeft: "10px",
              marginTop: "10px",
              backgroundColor: "white",
              border: "1px solid #ccc",
              fontSize: "16px",
              borderRadius: "2px",
              padding: "10px 10px",
              cursor: "pointer",
            }}
            onClick={sendToStreamlineAsComment}
          >
            Send
          </button>
          <button
            style={{
              marginLeft: "10px",
              marginTop: "10px",
              backgroundColor: "white",
              border: "1px solid #ccc",
              fontSize: "16px",
              borderRadius: "2px",
              padding: "10px 10px",
              cursor: "pointer",
            }}
            onClick={() => setIsOpen(false)}
          >
            Cancel
          </button>
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
