import * as React from "react";
import PropTypes from "prop-types";
import { makeStyles } from "@fluentui/react-components";
import { officeKeys, getOfficeKeyValue } from "../../config/utility.js";
import AuthorizationUI from "./AuthorizationUI";
import TokenUI from "./TokenUI";
import { refreshAccessToken } from "../../config/auth.js";
import { useDispatch } from "react-redux";
import { setTokenData } from "../../app/authSlice.js";
import { AlertMessageModal } from "./AlertMessageModal.jsx";
import { setAlertMessage, setLoading } from "../../app/loaderSlice.js";
import { Loader } from "./Loader.jsx";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = () => {
  const styles = useStyles();
  const [refreshToken, setRefreshToken] = React.useState("");
  const dispatchToRedux = useDispatch();
  React.useEffect(() => {
    const existingRefreshToken = getOfficeKeyValue(officeKeys.refreshToken);
    async function handleRefreshToken(refreshToken) {
      dispatchToRedux(setLoading(true));
      try {
        const newToken = await refreshAccessToken(refreshToken);
        dispatchToRedux(setTokenData(newToken));
        setRefreshToken(newToken.refreshToken);
      } catch (error) {
        dispatchToRedux(setAlertMessage({ message: error.message, intent: "error" }));
      }
      dispatchToRedux(setLoading(false));
    }
    if (existingRefreshToken) {
      handleRefreshToken(existingRefreshToken);
    }
  }, []);

  return (
    <div className={styles.root}>
      {!refreshToken ? <AuthorizationUI setRefreshToken={setRefreshToken} /> : <TokenUI />}
      <AlertMessageModal />
      <Loader />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
