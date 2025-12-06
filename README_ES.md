# 🚀 Automatización de WhatsApp Web en VB6

Este proyecto, desarrollado en **Visual Basic 6**, implementa una solución para automatizar el envío de mensajes a través de **WhatsApp Web** usando dos motores de navegación:

* 🧩 **RC6 + WebView2** (Proyecto de Olaf Schmidt)
* 🤖 **SeleniumBasic**

La aplicación permite elegir el motor deseado y ejecutar acciones de mensajería de forma integrada.

---

## ✨ Características principales

### **1. 💬 Envío de mensajes de texto**

Tanto **WebView2** como **SeleniumBasic** permiten enviar mensajes a contactos y grupos. Los mensajes pueden incluir:

* Texto
* Emojis 😃🔥🎉👌

### **2. 📁 Envío de archivos (solo con SeleniumBasic)**

Con **SeleniumBasic** es posible enviar:

* 🖼️ Imágenes
* 📄 Documentos
* 🎞️ Videos

> ⚠️ Esta función no está disponible con WebView2.

### **3. 🧱 Integración con RC6 (Olaf)**

Este proyecto usa componentes del framework **RC6** para trabajar con WebView2 en VB6, ofreciendo:

* Navegación moderna dentro del formulario
* Ejecución de JavaScript
* Manipulación de elementos HTML

---

## 🛠️ Tecnologías utilizadas

### **🧩 RC6 + WebView2**

* Basado en el trabajo de Olaf Schmidt
* Proporciona un navegador moderno dentro de VB6
* Permite enviar mensajes mediante JavaScript

### **🤖 SeleniumBasic**

* Automatización del navegador (Chrome/Edge)
* Acceso completo al DOM
* Envío de archivos y mensajes
* 🔄 **Actualización automática de WebDrivers**: El proyecto incluye una aplicación adicional desarrollada específicamente para actualizar los WebDrivers sin intervención manual. Esta herramienta gestiona la descarga, reemplazo y verificación de las versiones necesarias, garantizando que Selenium siempre opere con los controladores correctos.
* Automatización del navegador (Chrome/Edge)
* Acceso completo al DOM
* Envío de archivos y mensajes

---

## 🔍 Comparación de funcionalidades

| Función                  | WebView2 (RC6) | SeleniumBasic |
| ------------------------ | -------------- | ------------- |
| 💬 Envío de mensajes     | ✔️             | ✔️            |
| 😀 Envío de emojis       | ✔️             | ✔️            |
| 📁 Envío de archivos     | ❌              | ✔️            |
| 🧭 Navegación automática | ✔️             | ✔️            |
| 🔧 Control del DOM       | Parcial        | Completo      |

---

## 📦 Requisitos

Este proyecto **incluye las librerías RC6 y SeleniumBasic necesarias para su funcionamiento**, por lo que:

* No es necesario instalar dependencias de forma manual.
* Al ejecutar la aplicación, el proyecto se encarga de **registrar automáticamente** los componentes requeridos.

---

## 📦 Requisitos (detallado)

* **Visual Basic 6.0**
* **RC6 (con soporte WebView2)**
* **Microsoft WebView2 Runtime**
* **SeleniumBasic**
* Navegador compatible (Chrome/Edge)

---

## 🎯 Finalidad del proyecto

Herramienta orientada a automatizar procesos con WhatsApp Web, ideal para:

* 📢 Envío masivo de mensajes
* ⏰ Envío programado
* 🏛️ Integración con sistemas legacy en VB6

---

Si deseas colaborar, puedes abrir un **issue** o enviar un **pull request**. También puedes contactarme para solicitar permisos o resolver dudas sobre su uso.
