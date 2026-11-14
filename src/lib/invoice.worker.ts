import { createSample, generateArchive, parseWorkbook } from "./engine";
import type { WorkerRequest, WorkerResponse } from "./types";

const reply = (message: WorkerResponse, transfer: Transferable[] = []) =>
  self.postMessage(message, { transfer });
self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  try {
    const request = event.data;
    if (request.type === "parse")
      reply({ type: "parsed", result: await parseWorkbook(request.buffer) });
    if (request.type === "sample") {
      const buffer = await createSample();
      reply({ type: "sample", buffer }, [buffer]);
    }
    if (request.type === "generate") {
      const buffer = await generateArchive(
        request.template,
        request.invoices,
        (done, total) => reply({ type: "progress", done, total }),
      );
      reply({ type: "generated", buffer }, [buffer]);
    }
  } catch (error) {
    reply({
      type: "error",
      message:
        error instanceof Error
          ? error.message
          : "Đã xảy ra lỗi khi xử lý tệp Excel.",
    });
  }
};
