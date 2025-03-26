import { useState, useEffect } from "react";
import reactLogo from "./assets/react.svg";
import { invoke } from "@tauri-apps/api/core";
import { open } from '@tauri-apps/plugin-dialog';
import { listen, emit } from '@tauri-apps/api/event';
import "./App.css";

function App() {
  const [greetMsg, setGreetMsg] = useState("");
  const [name, setName] = useState("");
  const [response1, setResponse1] = useState("");
  const [response2, setResponse2] = useState("");
  const [response3, setResponse3] = useState("");
  const [file_info, setFileInfo] = useState("");
  // const [response4, setResponse4] = useState("");

  function open_dialog() {
    open({ multiple: false, directory: false }).then(files => {
      if (files && files.length > 0) {
        setFileInfo(`Selected file: ${files}`);
      } else {
        setFileInfo("No file selected");
      }
    });
  }

  async function greet() {
    // Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
    setGreetMsg(await invoke("greet", { name }));
  }

  function my_command() {
    invoke('my_command', { message: { field_str: 'some message', field_u32: 12 } }).then(message => {
      setResponse1(message.field_str); setResponse2(message.field_u32);
    });
  }

  function comm_with_error() {
    for (let arg of [1, 2]) {
      invoke('command_with_error', { arg })
        .then(message => { setResponse3(message); })
        .catch(message => { setResponse3(message); })
    }
  }

  function emitMessage() {
    emit('front-to-back', "hello from front")
  }

  useEffect(
    () => {
      let unlisten: any;
      async function f() {
        unlisten = await listen(
          'back-to-front',
          (event) => { console.log(`back-to-front ${event.payload} ${new Date()}`) }
        );
      }
      f();

      return () => {
        if (unlisten) {
          unlisten();
        }
      }
    },
    []
  )

  return (
    <main className="container">
      <h1>Welcome to Tauri + React</h1>

      <div className="row">
        <a href="https://vitejs.dev" target="_blank">
          <img src="/vite.svg" className="logo vite" alt="Vite logo" />
        </a>
        <a href="https://tauri.app" target="_blank">
          <img src="/tauri.svg" className="logo tauri" alt="Tauri logo" />
        </a>
        <a href="https://reactjs.org" target="_blank">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <p>Click on the Tauri, Vite, and React logos to learn more.</p>

      <form
        className="row"
        onSubmit={(e) => {
          e.preventDefault();
          greet();
        }}
      >
        <input
          id="greet-input"
          onChange={(e) => setName(e.currentTarget.value)}
          placeholder="Enter a name..."
        />
        <button type="submit">Greet</button>
      </form>
      <p>{greetMsg}</p>

      <form
        className="row"
        onSubmit={(e) => {
          e.preventDefault();
          my_command();
        }}
      >
        <button type="submit">コマンド実行</button>
      </form>
      <p>{response1}</p>
      <p>{response2}</p>

      <form
        className="row"
        onSubmit={(e) => {
          e.preventDefault();
          comm_with_error();
        }}
      >
        <button type="submit">エラーコマンド実行</button>
      </form>
      <p>{response3}</p>
      <button onClick={open_dialog}>Open Dialog</button>
      <p>{file_info}</p>
      <button onClick={emitMessage}>emit message on terminal</button>
    </main>
  );
}

export default App;
