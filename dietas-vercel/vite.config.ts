import {defineConfig} from "vite";
import react from "@vitejs/plugin-react";
import {fileURLToPath, URL} from "node:url";
export default defineConfig({build:{emptyOutDir:true},plugins:[react()],resolve:{alias:{"@":fileURLToPath(new URL(".",import.meta.url))}},server:{port:5173}});
