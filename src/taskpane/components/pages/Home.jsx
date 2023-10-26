import React from "react";
import HeroList from "../HeroList";
import TextInsertion from "../TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
/* global  */
const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

export default function Home() {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];
  return (
    <div className={styles.root}>
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      <TextInsertion />
    </div>
  );
}
