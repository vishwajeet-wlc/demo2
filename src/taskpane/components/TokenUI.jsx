/* global Office, fetch */
import React, { useState, useEffect } from "react";
import Dropdown from "./Dropdown.jsx";
import RequestStatus from "./RequestStatus.jsx";
import { setOfficeKeyValue, getOfficeKeyValue, officeKeys } from "../../config/utility.js";
import { useDispatch, useSelector } from "react-redux";
import { Text, Label, Button, makeStyles, Input } from "@fluentui/react-components";
import { setAlertMessage, setLoading } from "../../app/loaderSlice.js";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    minWidth: "100vw",
    position: "fixed",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    top: "0px",
    left: "0px",
    zIndex: "1000",
  },
  form: {
    width: "80%",
    display: "flex",
    flexDirection: "column",
  },
  input: {
    marginTop: "10px",
  },
  button: {
    marginTop: "20px",
    width: "50%",
  },
  title: {
    fontSize: "18px",
    marginBottom: "15px",
    marginTop: "10px",
  },
  label: {
    marginBottom: "2px",
    marginTop: "15px",
  },
});
function TokenUI() {
  const [clientToken, setClientToken] = useState("");
  const [clientDomain, setClientDomain] = useState("");
  const [auth, setAuth] = useState(false);
  const [formTypes, setFormTypes] = useState([]);
  const [reqId, setReqId] = useState("");
  const authData = useSelector((state) => state.auth);
  const classes = useStyles();
  const dispatchToRedux = useDispatch();

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
      return dispatchToRedux(
        setAlertMessage({ message: "client token and domain both are required", intent: "warning" })
      );
    }

    try {
      dispatchToRedux(setLoading(true));
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
        return dispatchToRedux(setAlertMessage({ message: "Invalid Client Token.", intent: "error" }));
      }

      if (response.status === 401) {
        return dispatchToRedux(
          setAlertMessage({
            message: "Please sign in to Streamline using your Google account to continue.",
            intent: "error",
          })
        );
      }

      if (response.status !== 204) {
        return dispatchToRedux(
          setAlertMessage({ message: "Please check the client token and domain.", intent: "error" })
        );
      }

      setOfficeKeyValue(officeKeys.clientToken, clientToken);
      setOfficeKeyValue(officeKeys.clientDomain, clientDomain);
      await fetchAndSaveStreamlineForms(clientDomain, clientToken.split(".").pop());
      setAuth(true);
      dispatchToRedux(setLoading(false));
      dispatchToRedux(setAlertMessage({ message: "Saved successfully.", intent: "success" }));
    } catch (error) {
      dispatchToRedux(setLoading(false));
      return dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
    }
  }

  async function fetchAndSaveStreamlineForms(clientDomain, orgId) {
    try {
      dispatchToRedux(setLoading(true));
      const response = await fetch(`${clientDomain}/api/outlook/request-forms/${orgId}`, {
        headers: {
          Authorization: `Bearer ${authData.accessToken}`,
        },
      });
      const formData = await response.json();
      setFormTypes([...formData]);

      if (response.status === 400) {
        return dispatchToRedux(
          setAlertMessage({
            message: "Ensure integration is enabled and log in with your microsoft account in Streamline.",
            intent: "error",
          })
        );
      }

      if (response.status === 401) {
        return dispatchToRedux(
          setAlertMessage({
            message: "Please sign in to Streamline using your Google account to continue.",
            intent: "error",
          })
        );
      }
      dispatchToRedux(setLoading(false));
      dispatchToRedux(setAlertMessage({ message: "Streamline form fetched successfully.", intent: "success" }));
    } catch (error) {
      dispatchToRedux(setLoading(false));
      return dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
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
          <div className={classes.root}>
            <Text className={classes.title} variant="large" weight="bold">
              Setup access with streamline
            </Text>
            <form className={classes.form}>
              <Label className={classes.label} htmlFor="client-token">
                {" "}
                Client Token{" "}
              </Label>
              <Input
                id="client-token"
                placeholder="Client Token"
                className={classes.input}
                onChange={(e) => {
                  setClientToken(e.target.value);
                }}
              />
              <Label className={classes.label} htmlFor="domain-name">
                Domain Name
              </Label>
              <Input
                id="domain-name"
                placeholder="Streamline Domain"
                className={classes.input}
                onChange={(e) => {
                  setClientDomain(e.target.value);
                }}
              />
              <Button appearance="primary" className={classes.button} onClick={saveStreamlineSettings}>
                Save Settings
              </Button>
            </form>
          </div>
        </>
      )}
    </>
  );
}

TokenUI.propTypes = {};

export default TokenUI;
