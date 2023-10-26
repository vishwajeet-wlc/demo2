import React, { useState } from "react";
import MainForm from "./MainForm";
import PropTypes from "prop-types";

function Dropdown({ formDetails, clientDomain, orgId }) {
  const [selectedOption, setSelectedOption] = useState("");
  const [selectedOptionObject, setSelectedOptionObject] = useState({});

  const handleOptionChange = (event) => {
    setSelectedOption(event.target.value);

    const form = formDetails.find((res) => {
      return res.matter === event.target.value;
    });
    setSelectedOptionObject(form);
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
        <div className="dropdown-container">
          <select
            style={{
              width: "90%",
              border: "1px solid #eee",
              padding: "10px 10px",
              marginLeft: "10px",
              borderRadius: "2px",
              fontSize: "18px",
            }}
            id="dropdown"
            value={selectedOption}
            onChange={handleOptionChange}
          >
            <option value="">Choose a form </option>

            {formDetails.map((option) => (
              <option key={option._id} value={option.matter}>
                {option.matter}
              </option>
            ))}
          </select>
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
