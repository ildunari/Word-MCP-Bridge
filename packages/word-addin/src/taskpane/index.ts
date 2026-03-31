import { mount } from "svelte";
import App from "./components/app.svelte";
import "./index.css";

Office.onReady(() => {
  const target = document.getElementById("container");
  if (!target) return;

  mount(App, { target });
});
