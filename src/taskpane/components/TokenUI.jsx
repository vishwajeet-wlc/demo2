import React, { useState, useEffect } from "react";
import Dropdown from "./Dropdown.jsx";
import RequestStatus from "./RequestStatus.jsx";
import { setOfficeKeyValue, getOfficeKeyValue, officeKeys } from "../../config/utility.js";
import { useSelector } from "react-redux";
/* global Office, alert, fetch */

function TokenUI() {
  const [clientToken, setClientToken] = useState("");
  const [clientDomain, setClientDomain] = useState("");
  const [auth, setAuth] = useState(false);
  const [formTypes, setFormTypes] = useState([]);
  const [reqId, setReqId] = useState("");
  const authData = useSelector((state) => state.auth);

  useEffect(() => {
    const conversation = Office.context?.mailbox.initialData.conversationId;
    const requestId = getOfficeKeyValue(conversation);
    if (requestId) {
      setReqId(requestId);
    }
    const domain = getOfficeKeyValue(officeKeys.clientDomain);
    const token = getOfficeKeyValue(officeKeys.clientToken);

    if (domain && token) {
      setAuth(true);
      const orgId = token.split(".").pop();
      fetchAndSaveStreamlineForms(domain, orgId);
      setClientToken(token);
      setClientDomain(domain);
    }
  }, []);

  async function saveStreamlineSettings(e) {
    e.preventDefault();
    if (!clientToken || !clientDomain) {
      return alert("Failed: Client token and domain both are required");
    }

    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${authData.accessToken}`,
      },
      body: JSON.stringify({
        clientToken,
        clientDomain,
      }),
    };
    const response = await fetch(`${clientDomain}/api/outlook/validate`, options);

    if (response.status === 400) {
      return alert("Failed: Invalid Client Token.");
    }

    if (response.status === 401) {
      return alert("Failed: Please sign in to Streamline using your Google account to continue.");
    }

    if (response.status !== 204) {
      return alert("Failed: Please check the client token and domain.");
    }

    setOfficeKeyValue(officeKeys.clientToken, clientToken);
    setOfficeKeyValue(officeKeys.clientDomain, clientDomain);
    await fetchAndSaveStreamlineForms(clientDomain, clientToken.split(".").pop());
    setAuth(true);
  }

  async function fetchAndSaveStreamlineForms(clientDomain, orgId) {
    const response = await fetch(`${clientDomain}/api/outlook/request-forms/${orgId}`, {
      headers: {
        Authorization: `Bearer ${authData.accessToken}`,
      },
    });
    const formData = await response.json();
    setFormTypes([...formData]);
    if (response.status === 400) {
      return alert("Failed: Ensure integration is enabled and log in with your Google account in Streamline.");
    }
  }

  return (
    <>
      {reqId && auth ? (
        <RequestStatus reqId={reqId} clientToken={clientToken} clientDomain={clientDomain} />
      ) : auth ? (
        <>
          <Dropdown formDetails={formTypes} clientDomain={clientDomain} orgId={clientToken.split(".").pop()} />
        </>
      ) : (
        <>
          <p
            style={{
              fontWeight: "bold",
              fontFamily: "sans-serif",
              fontSize: "20px",
              padding: "10px",
              border: "1px solid #ccc",
              width: "100%",
            }}
          >
            Setup access with streamline
          </p>
          <form>
            <label
              style={{
                marginTop: "10px",
                paddingLeft: "5%",
                fontSize: "18px",
                width: "80%",
                fontFamily: "sans-serif",
              }}
            >
              Client Token
            </label>
            <br />
            <input
              placeholder="Client Token"
              style={{
                marginTop: "10px",
                marginLeft: "5%",
                marginRight: "5%",
                width: "80%",
                border: "1px solid #eee",
                borderRadius: "2px",
                fontSize: "14px",
                padding: "10px 10px",
              }}
              onChange={(e) => {
                setClientToken(e.target.value);
              }}
            />{" "}
            <br />
            <div style={{ marginTop: "20px" }}>
              <label
                style={{
                  marginTop: "10px",
                  paddingLeft: "5%",
                  fontSize: "18px",
                  width: "80%",
                  fontFamily: "sans-serif",
                }}
              >
                Domain Name
              </label>{" "}
              <br />
              <input
                placeholder="Streamline Domain"
                onChange={(e) => {
                  setClientDomain(e.target.value);
                }}
                style={{
                  marginTop: "10px",
                  marginLeft: "5%",
                  marginRight: "5%",
                  width: "80%",
                  border: "1px solid #eee",
                  borderRadius: "2px",
                  fontSize: "14px",
                  padding: "10px",
                }}
              />
            </div>
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
              onClick={(e) => {
                saveStreamlineSettings(e);
              }}
            >
              Save Settings
            </button>
          </form>
        </>
      )}
    </>
  );
}

TokenUI.propTypes = {};

export default TokenUI;
