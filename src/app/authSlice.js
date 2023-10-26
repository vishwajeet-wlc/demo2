import { createSlice } from "@reduxjs/toolkit";

const initialState = {
  accessToken: "",
  refreshToken: "",
  expiresAt: 0,
};

const authSlice = createSlice({
  name: "auth",
  initialState,
  reducers: {
    setTokenData: (state, { payload }) => {
      state.accessToken = payload.accessToken;
      state.refreshToken = payload.refreshToken;
      state.expiresAt = payload.expiresAt;
    },
  },
});

const { setTokenData } = authSlice.actions;

export { setTokenData };
export default authSlice.reducer;
