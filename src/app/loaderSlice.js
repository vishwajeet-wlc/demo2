import { createSlice } from "@reduxjs/toolkit";

const initialState = {
  isLoading: false,
  alertMessage: "",
  intent: "",
};

const loaderSlice = createSlice({
  name: "loader",
  initialState,
  reducers: {
    setLoading: (state, { payload }) => {
      state.isLoading = payload;
    },
    setAlertMessage: (state, { payload }) => {
      state.alertMessage = payload.message;
      state.intent = payload.intent;
    },
  },
});

const { setLoading, setAlertMessage } = loaderSlice.actions;

export { setLoading, setAlertMessage };
export default loaderSlice.reducer;
