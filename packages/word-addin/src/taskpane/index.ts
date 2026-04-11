import { mount } from "svelte";
import "../commands/commands";
import { ensureWordTaskpaneRuntime } from "../lib/shared-runtime";
import App from "./components/app.svelte";
import "./index.css";

Office.onReady(() => {
  void ensureWordTaskpaneRuntime();
  const target = document.getElementById("container");
  if (!target) return;

  mount(App, { target });
});
