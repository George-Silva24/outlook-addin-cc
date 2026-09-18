/*
  ALTERA AQUI: o(s) email(s) que queres garantir que ficam em CC.
*/
const EMAIL_OBRIGATORIO = "exemplo@empresa.com";

Office.onReady(() => {
  // Necessário para o runtime de eventos, mesmo sem UI.
});

// Função chamada automaticamente pelo Outlook antes de o email ser enviado.
function checkCC(event) {
  const item = Office.context.mailbox.item;

  item.cc.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      // Se não conseguirmos ler o CC, deixamos o envio seguir em frente
      // em vez de bloquear o utilizador por um erro técnico.
      event.completed({ allowEvent: true });
      return;
    }

    const ccRecipients = result.value || [];
    const jaTemEmail = ccRecipients.some(
      (r) => r.emailAddress.toLowerCase() === EMAIL_OBRIGATORIO.toLowerCase()
    );

    if (jaTemEmail) {
      // Já está em CC — envia normalmente, sem perguntar nada.
      event.completed({ allowEvent: true });
      return;
    }

    // Não está em CC — bloqueia o envio e pergunta ao utilizador.
    Office.context.mailbox.item.notificationMessages.addAsync("ccWarning", {
      type: "insightMessage",
      message: `O email ${EMAIL_OBRIGATORIO} não está em CC. Adiciona-o e envia novamente, ou usa a caixa de diálogo para decidir agora.`,
      icon: "icon1",
      persistent: false
    });

    // Abre uma caixa de diálogo a perguntar se quer adicionar o email.
    Office.context.ui.displayDialogAsync(
      "https://SEU-DOMINIO-OU-GITHUB-PAGES-AQUI/dialog.html",
      { height: 25, width: 25, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          // Se o diálogo não abrir, bloqueia o envio por segurança.
          event.completed({ allowEvent: false, errorMessage: "Não foi possível abrir a caixa de diálogo." });
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          const resposta = arg.message; // "sim" ou "nao", enviado pelo dialog.html

          if (resposta === "sim") {
            item.cc.addAsync([EMAIL_OBRIGATORIO], () => {
              dialog.close();
              event.completed({ allowEvent: true });
            });
          } else {
            dialog.close();
            event.completed({ allowEvent: true }); // deixa enviar sem o CC
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          // Utilizador fechou o diálogo sem responder — bloqueia por segurança.
          event.completed({ allowEvent: false });
        });
      }
    );
  });
}

// Regista a função para o runtime de eventos conseguir encontrá-la.
Office.actions.associate("checkCC", checkCC);
