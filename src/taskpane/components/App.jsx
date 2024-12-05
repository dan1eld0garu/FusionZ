import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import FusionZPane from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { processDocument, clearEdits } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.

  return (
    <div className={styles.root}>
      <Header logo="assets/Fusionz_logo_symbol.png" title={title} message="FusionZ" />
      {/* <HeroList message="Discover what this add-in can do for you today!" items={listItems} /> */}
      <FusionZPane processDocument={processDocument} clearEdits={clearEdits} />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
