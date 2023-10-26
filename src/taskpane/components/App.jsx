import * as React from "react";
import PropTypes from "prop-types";
// import Header from "./Header";
// import HeroList from "./HeroList";
// import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
// import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { officeKeys, getOfficeKeyValue } from "../../config/utility.js";
import AuthorizationUI from "./AuthorizationUI";
import TokenUI from "./TokenUI";
import { refreshAccessToken } from "../../config/auth.js";
import { useDispatch } from "react-redux";
import { setTokenData } from "../../app/authSlice.js";

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
    async function handleRefreshToken() {
      const newToken = await refreshAccessToken("6f067864-af43-4c81-be6b-cd09a97314d3", existingRefreshToken);
      dispatchToRedux(setTokenData(newToken));
      setRefreshToken(newToken.refreshToken);
    }
    if (existingRefreshToken) {
      handleRefreshToken();
    }
  }, []);

  return (
    <div className={styles.root}>
      {!refreshToken ? <AuthorizationUI setRefreshToken={setRefreshToken} /> : <TokenUI />}
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
