import React, { useState } from "react";
import DynamicForm from "./DynamicForm";
import RequestStatus from "./RequestStatus";
import PropTypes from "prop-types";

function MainForm({ selectedFormDetails, domain, orgId }) {
  const [reqId, setReqId] = useState("");
  return (
    <>
      <div>
        {reqId ? (
          <RequestStatus reqId={reqId} clientToken={`text.${orgId}`} clientDomain={domain} />
        ) : (
          <DynamicForm
            selectedFormDetails={selectedFormDetails}
            domain={domain}
            orgId={orgId}
            setRequestId={setReqId}
          />
        )}
      </div>
    </>
  );
}
MainForm.propTypes = {
  selectedFormDetails: PropTypes.object.isRequired,
  domain: PropTypes.string.isRequired,
  orgId: PropTypes.string.isRequired,
};
export default MainForm;
