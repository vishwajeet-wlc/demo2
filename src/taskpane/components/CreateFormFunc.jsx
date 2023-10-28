import React from "react";
import PropTypes from "prop-types";
import { Checkbox, Input, Radio, Select, Text, makeStyles } from "@fluentui/react-components";

// Define styling, split out styles for each area.
const useStyles = makeStyles({
  container: {
    marginTop: "20px",
    marginLeft: "10px",
    display: "flex",
    flexDirection: "column",
  },
  text: {
    fontSize: "16px",
    width: "90%",
    fontFamily: "sans-serif",
    marginBottom: "5px",
  },
  input: {
    width: "80%",
    fontSize: "14px",
  },
  select: {
    marginRight: "5%",
    width: "80%",
    fontSize: "14px",
  },
});

export default function CreateWidgetForCardUI({ field, onChange, values, attachments }) {
  const classes = useStyles();
  switch (field.type) {
    case "text":
    case "paragraph":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <Input
            type="text"
            className={classes.input}
            placeholder={field.title}
            name={field._id}
            onChange={onChange}
            value={values[field._id] ?? ""}
          />
        </div>
      );

    case "number":
    case "currency":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <Input
            type="number"
            className={classes.input}
            placeholder={field.title}
            name={field._id}
            onChange={onChange}
            value={values[field._id] ?? ""}
          />
        </div>
      );

    case "date":
    case "futuredate":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <Input
            type="date"
            className={classes.input}
            placeholder={field.title}
            name={field._id}
            onChange={onChange}
            value={values[field._id] ?? ""}
          />
        </div>
      );

    case "email":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <Input
            type="email"
            className={classes.input}
            placeholder={field.title}
            name={field._id}
            onChange={onChange}
            value={values[field._id] ?? ""}
          />
        </div>
      );

    case "radio":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <div>
            {field.options.map((option) => {
              return (
                <Radio
                  label={option.value}
                  key={option._id}
                  name={field._id}
                  onChange={onChange}
                  value={option.value}
                />
              );
            })}
          </div>
        </div>
      );

    case "checkbox":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <div>
            {field.options.map((option) => {
              return (
                <Checkbox
                  label={option.value}
                  key={option._id}
                  name={field._id}
                  onChange={onChange}
                  value={option.value}
                  as="input"
                />
              );
            })}
          </div>
        </div>
      );

    case "select":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <Select className={classes.select} name={field._id} placeholder={field.title} onChange={onChange}>
            <option value=""> Select value</option>
            {field.options.map((opt) => {
              return (
                <option key={opt._id} value={opt.value}>
                  {opt.value}
                </option>
              );
            })}
          </Select>
        </div>
      );
    case "attachment":
      return (
        <div className={classes.container}>
          <Text className={classes.text}>{field.title}</Text>
          <Select className={classes.select} name="attachment" placeholder={"Select an attachment"} onChange={onChange}>
            {!attachments.length ? <option> No Attachments </option> : <option> Select Attachment </option>}
            {attachments.length &&
              attachments.map((details) => {
                return (
                  <option key={details.id} value={details.id}>
                    {details.name}
                  </option>
                );
              })}
          </Select>
        </div>
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
};
