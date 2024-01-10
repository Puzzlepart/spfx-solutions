import { useState } from "react";

export default function useLocationHash(): string {
    const [hash, setHash] = useState(window.location.hash);
    window.addEventListener("hashchange", () => setHash(window.location.hash));
  
    return hash.substring(1);
  }