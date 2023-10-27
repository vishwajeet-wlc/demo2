import React, { useState } from "react";
import MainForm from "./MainForm";
import PropTypes from "prop-types";
import { useDispatch } from "react-redux";
import { setAlertMessage } from "../../app/loaderSlice";
import { Select, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    width: "90%",
    marginLeft: "10px",
  },
  select: {
    fontSize: "18px",
  },
});
function Dropdown({ formDetails, clientDomain, orgId }) {
  const [selectedOption, setSelectedOption] = useState("");
  const [selectedOptionObject, setSelectedOptionObject] = useState({});
  const dispatchToRedux = useDispatch();
  const classes = useStyles();

  const handleOptionChange = (event) => {
    setSelectedOption(event.target.value);

    const form = formDetails.find((res) => {
      return res.matter === event.target.value;
    });
    setSelectedOptionObject(form);
    dispatchToRedux(setAlertMessage({ message: "Form Switched successfully", intent: "success" }));
  };

  return (
    <>
      {selectedOption ? (
        <>
          <MainForm
            formName={selectedOption}
            allForms={formDetails}
            selectedFormDetails={selectedOptionObject}
            domain={clientDomain}
            orgId={orgId}
          />
        </>
      ) : (
        <div className={classes.container}>
          <Select id="dropdown" value={selectedOption} onChange={handleOptionChange} className={classes.select}>
            <option value="">Choose a form</option>
            {formDetails.map((option) => (
              <option key={option._id} value={option.matter}>
                {option.matter}
              </option>
            ))}
          </Select>
        </div>
      )}
    </>
  );
}

Dropdown.propTypes = {
  formDetails: PropTypes.array.isRequired,
  clientDomain: PropTypes.string.isRequired,
  orgId: PropTypes.string.isRequired,
};
export default Dropdown;
