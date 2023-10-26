import React, { useEffect } from "react";
import { Button, Image, tokens, makeStyles } from "@fluentui/react-components";
import logo from "/assets/logo-filled.png";
import { getMicrosoftAccessToken } from "../../config/auth";
import PropTypes from "prop-types";
import { useDispatch } from "react-redux";
import { setTokenData } from "../../app/authSlice";
import { setOfficeKeyValue, officeKeys } from "../../config/utility";
/* global window */
const useStyles = makeStyles({
  root: {
    minHeight: "75vh",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingBottom: "30px",
    paddingTop: "100px",
    gap: "30px",
    backgroundColor: tokens.colorNeutralBackground3,
  },
});

export default function AuthorizationUI({ setRefreshToken }) {
  const styles = useStyles();
  const dispatchToRedux = useDispatch();

  async function handleGetAccessToken() {
    try {
      const tokenData = await getMicrosoftAccessToken("6f067864-af43-4c81-be6b-cd09a97314d3");
      if (tokenData) {
        setOfficeKeyValue(officeKeys.refreshToken, tokenData.refreshToken);
        setRefreshToken(tokenData.refreshToken);
        dispatchToRedux(setTokenData(tokenData));
      }
    } catch (error) {
      console.log(error);
    }
  }

  useEffect(() => {
    const oauthParams = Object.fromEntries(new URLSearchParams(window.location.search));
    if (oauthParams.code) {
      window.opener.postMessage(true, window.location.origin);
      window.close();
    }
  }, [window.location.search]);

  return (
    <div className={styles.root}>
      <Image width="90" height="90" src={logo} alt={"streamline-ai-logo"} />
      <Button appearance="primary" disabled={false} size="large" onClick={handleGetAccessToken}>
        {" "}
        Authorize Addon{" "}
      </Button>
    </div>
  );
}

AuthorizationUI.propTypes = {
  setRefreshToken: PropTypes.func.isRequired,
};
