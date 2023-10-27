/* global setTimeout */
import * as React from "react";
import { useId, Toaster, useToastController, Toast, ToastTitle, ToastBody, Button } from "@fluentui/react-components";
import { useDispatch, useSelector } from "react-redux";
import { setAlertMessage } from "../../app/loaderSlice";
import { ArrowExitRegular } from "@fluentui/react-icons";

export const AlertMessageModal = () => {
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const loaderData = useSelector((state) => state.loader);
  const dispatchToRedux = useDispatch();

  function handleClose() {
    dispatchToRedux(setAlertMessage({ message: "", intent: "" }));
  }
  const notify = () =>
    dispatchToast(
      <Toast>
        <ToastTitle action={<Button onClick={handleClose} icon={<ArrowExitRegular />} />}>
          {" "}
          {loaderData?.intent}{" "}
        </ToastTitle>
        <ToastBody> {loaderData.alertMessage} </ToastBody>
      </Toast>,
      { intent: loaderData.intent, timeout: 10000 }
    );

  React.useEffect(() => {
    if (loaderData.alertMessage) {
      notify();
      setTimeout(() => {
        dispatchToRedux(setAlertMessage({ message: "", intent: "" }));
      }, 10000);
    }
  }, [loaderData]);

  return <>{loaderData.alertMessage && <Toaster toasterId={toasterId} />}</>;
};
