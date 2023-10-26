import React from "react";
import PropTypes from "prop-types";

export default function CreateWidgetForCardUI({ field, onChange, values, attachments, getAttachmentData }) {
  switch (field.type) {
    case "text":
    case "paragraph":
      return (
        <input
          type="text"
          style={{
            width: "80%",
            border: "1px solid #ccc",
            borderRadius: "2px",
            fontSize: "14px",
            padding: "10px 10px",
          }}
          placeholder={field.title}
          name={field._id}
          onChange={onChange}
          value={values[field._id]}
        />
      );

    case "number":
    case "currency":
      return (
        <input
          type="number"
          style={{
            marginRight: "5%",
            width: "80%",
            border: "1px solid #ccc",
            borderRadius: "2px",
            fontSize: "14px",
            padding: "10px 10px",
          }}
          placeholder={field.title}
          name={field._id}
          onChange={onChange}
          value={values[field._id]}
        />
      );

    case "date":
    case "futuredate":
      return (
        <input
          type="date"
          style={{
            marginRight: "5%",
            width: "80%",
            border: "1px solid #ccc",
            borderRadius: "2px",
            fontSize: "14px",
            padding: "10px 10px",
          }}
          placeholder={field.title}
          name={field._id}
          onChange={onChange}
          value={values[field._id]}
        />
      );

    case "email":
      return (
        <input
          type="email"
          style={{
            marginRight: "5%",
            width: "80%",
            border: "1px solid #ccc",
            borderRadius: "2px",
            fontSize: "14px",
            padding: "10px 10px",
          }}
          placeholder={field.title}
          name={field._id}
          onChange={onChange}
          value={values[field._id]}
        />
      );

    case "radio":
      return (
        <div>
          {field.options.map((option) => {
            return <input key={option.value} type="radio" name={field._id} onChange={onChange} value={option.value} />;
          })}
        </div>
      );

    case "checkbox":
      return (
        <div>
          {field.options.map((option) => {
            return (
              <input key={option.value} type="checkbox" name={field._id} onChange={onChange} value={option.value} />
            );
          })}
        </div>
      );

    case "select":
      return (
        <select
          style={{
            marginRight: "5%",
            width: "90%",
            border: "1px solid #ccc",
            borderRadius: "2px",
            fontSize: "14px",
            padding: "10px 10px",
          }}
          name={field._id}
          placeholder={field.title}
          onChange={onChange}
        >
          <option value=""> Select </option>
          {field.options.map((opt) => {
            return (
              <option key={opt.value} value={opt.value}>
                {opt.value}
              </option>
            );
          })}
        </select>
      );
    case "attachment":
      return (
        <select
          style={{
            marginRight: "5%",
            width: "90%",
            border: "1px solid #ccc",
            borderRadius: "2px",
            fontSize: "14px",
            padding: "10px 10px",
          }}
          name={field._id}
          placeholder={"Select an attachment"}
          onChange={onChange}
        >
          {!attachments.length ? <option> No Attachments </option> : <option> Select Attachment </option>}
          {attachments.length &&
            attachments.map((details) => {
              getAttachmentData(details);
              return <option key={details.name}>{details.name}</option>;
            })}
        </select>
      );
    default:
      return null;
  }
}

CreateWidgetForCardUI.propTypes = {
  field: PropTypes.object,
  onChange: PropTypes.func.isRequired,
  values: PropTypes.object.isRequired,
  attachments: PropTypes.array.isRequired,
  getAttachmentData: PropTypes.func.isRequired,
};
