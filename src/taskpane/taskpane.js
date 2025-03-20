Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
        document.getElementById("uploadAttachmentButton").addEventListener("click", sendEmailAttachmentToServer);
        document.getElementById("authenticateUser").addEventListener("click", authenticateUser);

    }
});



document.addEventListener("DOMContentLoaded", function() {
    const loginModal = document.getElementById("loginModal");
    const openLoginModalBtn = document.getElementById("openLoginModal");
    const closeLoginModalBtn = document.querySelector(".close");
    const authenticateUserBtn = document.getElementById("authenticateUser");

    openLoginModalBtn.onclick = function() {
        loginModal.style.display = "block";
    };

    closeLoginModalBtn.onclick = function() {
        loginModal.style.display = "none";
    };

    window.onclick = function(event) {
        if (event.target === loginModal) {
            loginModal.style.display = "none";
        }
    };

    authenticateUserBtn.onclick = authenticateUser;

});

async function run() {
    const item = Office.context.mailbox.item;
    let insertAt = document.getElementById("item-subject");
    let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
    insertAt.appendChild(label);
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(item.subject));
    insertAt.appendChild(document.createElement("br"));
}

async function authenticateUser() {
    const username = document.getElementById("username").value;
    const password = document.getElementById("password").value;
    const serverUrl = localStorage.getItem("serverUrl");

    if (!username || !password || !serverUrl) {
        document.getElementById("authenticateStatus").innerText = " Please enter username, password, and server URL.";
        return;
    }

    globalUsername = username;
    globalPassword = password;
    globalServerUrl = serverUrl;
    
    // ✅ Sauvegarde des informations d'identification en local
    localStorage.setItem("username", username);
    localStorage.setItem("password", password);
    localStorage.setItem("serverUrl", serverUrl);
    
    document.getElementById("usernameDisplay").innerText = username;
    document.getElementById("authenticateStatus").innerText = "✅ Authentification réussie!";
    setTimeout(() => {
        document.getElementById("loginModal").style.display = "none"; 
    }, 1000);
}

async function sendEmailAttachmentToServer() {
    const item = Office.context.mailbox.item;
    const statusMessage = document.getElementById("uploadStatus");

    if (!globalUsername || !globalPassword ) {
    
        statusMessage.innerText = " Vous devez vous connecter avant d'envoyer un document ! ";
        return;
    }

    if (!item.attachments || item.attachments.length === 0) {
        console.warn(" Aucune pièce jointe trouvée.");
        statusMessage.innerText = " Aucune pièce jointe trouvée dans l'email.";
        return;
    }

    statusMessage.innerText = "⏳ Envoi en cours...";

    for (let attachment of item.attachments) {
        if (attachment.isInline) {
            console.log(` ${attachment.name} est une image inline, ignorée.`);
            continue;
        }

        console.log(` Récupération de la pièce jointe : ${attachment.name}`);

        item.getAttachmentContentAsync(attachment.id, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let fileData = result.value.content;
                let mimeType = result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64
                    ? "application/pdf"
                    : result.value.format;

                const fileBlob = new Blob([_base64ToArrayBuffer(fileData)], { type: mimeType });

                const metadata = {
                    name: attachment.name,
                    doctype: "fct",
                    mimetype: mimeType
                };
                const metadataBlob = new Blob([JSON.stringify(metadata)], { type: "application/json" });

                const formData = new FormData();
                formData.append("jsondata", metadataBlob);
                formData.append("document", fileBlob, attachment.name);

                console.log("📡 Envoi au serveur...");
       
                try {
                    const response = await fetch("scopsoftware/api/scopmaster/piece/addAsync", {

                        method: "POST",
                        headers: {
                            Authorization: `Basic ${btoa(globalUsername + ":" + globalPassword)}`
                        },
                        body: formData
                    });

                    if (!response.ok) {
                        throw new Error(` Erreur HTTP : ${response.status}`);
                    }

                    const result = await response.json();
                    console.log(" Réponse du serveur :", result);
                    statusMessage.innerText = "✅ Pièce jointe envoyée avec succès !";
                } catch (error) {
                    console.error(" Erreur lors de l'envoi :", error);
                    statusMessage.innerText = " Erreur lors de l'envoi de la pièce jointe.";
                }
            } else {
                console.error(" Erreur lors de la récupération de la pièce jointe :", result.error.message);
                statusMessage.innerText = " Impossible de récupérer la pièce jointe.";
            }
        });
    }
}

// Fonction utilitaire pour convertir Base64 en ArrayBuffer
function _base64ToArrayBuffer(base64) {
    let binaryString = atob(base64);
    let bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}