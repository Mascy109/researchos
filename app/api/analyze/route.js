export const dynamic = "force-dynamic";

import OpenAI from "openai";

export async function POST(request) {
  try {
    const body = await request.json();
    const transcript = body.transcript;

    if (!transcript) {
      return Response.json(
        { error: "No transcript provided." },
        { status: 400 }
      );
    }

    const openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    const response = await openai.responses.create({
      model: "gpt-5.6",
      input: [
        {
          role: "system",
          content:
            "You are a qualitative market research analyst. Analyze the supplied transcript carefully. Identify the most important research insight and support it only with evidence from the supplied transcript.",
        },
        {
          role: "user",
          content: `Analyze this transcript:

${transcript}

Return:
1. Main insight
2. Evidence
3. Implication`,
        },
      ],
    });

    return Response.json({
      result: response.output_text,
    });
  } catch (error) {
    console.error(error);

    return Response.json(
      {
        error: error.message || "AI analysis failed.",
      },
      { status: 500 }
    );
  }
}
