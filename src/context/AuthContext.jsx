import React, { useState } from "react";
import { createContext } from "react";
import PropTypes from "prop-types";

const AuthContext = createContext(null);

export default function AuthContextProvider({ children }) {
  const [token, setToken] = useState({});
  return <AuthContext.Provider value={{ token, setToken }}>{children}</AuthContext.Provider>;
}

AuthContextProvider.propTypes = {
  children: PropTypes.element,
};
