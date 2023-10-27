import * as React from "react";
import { Spinner, makeStyles } from "@fluentui/react-components";
import { useSelector } from "react-redux";

const useStyles = makeStyles({
  root: {
    backgroundColor: "#FFF",
    minHeight: "100vh",
    minWidth: "100vw",
    position: "fixed",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    top: "0px",
    left: "0px",
    zIndex: "1000",
  },
});

export const Loader = (props) => {
  const classes = useStyles();
  const loaderData = useSelector((state) => state.loader);

  return (
    loaderData.isLoading && (
      <div className={classes.root}>
        <Spinner appearance="primary" {...props} />
      </div>
    )
  );
};
