/**
 * Vercel Serverless 入口：将全站流量交给 Express（见 vercel.json rewrites）
 * 勿在本文件重复 listen；server.js 在 VERCEL=1 时不会监听端口。
 */
import app from '../server.js';

export default app;
