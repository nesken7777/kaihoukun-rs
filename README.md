# kaihoukun-rs
開放くんのリバースエンジニアリング

## 概要

これは[開放くん](http://cres.s28.xrea.com/soft/kaihoukun.html)をリバースエンジニアリングし、それをほぼそのままRustに転写したものです。

## ビルド方法
Rustの環境が構築されていれば、
```cmd
cargo build
```
のみで十分です。

**32ビットアプリケーションとしてビルドを行いたい場合:** ビルドターゲットを([ここ](https://doc.rust-lang.org/rustc/platform-support.html)や[ここ](https://doc.rust-lang.org/cargo/reference/config.html)を参考に)`i686-pc-windows-msvc`に変更してください。また、`app.manifest`の`amd64`と書かれている部分(2カ所)を`X86`に変更してください。
