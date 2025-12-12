// OpenAI 서비스 레이어
import OpenAI from "openai";
import dotenv from "dotenv";

// 환경변수 로드 (.env 파일에서)
dotenv.config();

// OpenAI API 키 검증
const apiKey = process.env.OPENAI_API_KEY;
if (!apiKey || apiKey === "your-api-key-here" || apiKey.trim() === "") {
  console.warn("⚠️  경고: OPENAI_API_KEY 환경변수가 설정되지 않았거나 유효하지 않습니다.");
  console.warn("   .env 파일을 생성하고 OPENAI_API_KEY를 설정해주세요.");
  console.warn("   .env.example 파일을 참고하세요.");
}

// OpenAI 클라이언트 초기화
// 환경변수에서만 API 키를 가져옴 (보안)
const openai = (apiKey && apiKey !== "your-api-key-here" && apiKey.trim() !== "") 
  ? new OpenAI({
      apiKey: apiKey.trim(),
    }) 
  : null;

if (openai) {
  console.log("✅ OpenAI 클라이언트 초기화 완료");
} else {
  console.log("⚠️  OpenAI 클라이언트 초기화 실패 - 기존 방식(string-similarity)으로 매칭합니다.");
}

/**
 * OpenAI를 사용하여 텍스트 매칭 개선
 * @param {string} text1 - 첫 번째 텍스트
 * @param {string[]} textList - 비교할 텍스트 목록
 * @returns {Promise<Object>} 매칭 결과
 */
export async function improveMatchingWithAI(text1, textList) {
  if (!openai) {
    console.error("❌ OpenAI 클라이언트가 초기화되지 않았습니다. API 키를 확인해주세요.");
    return null;
  }
  
  try {
    // 전체 목록 사용 (2024 시트의 모든 적요와 비교)
    const 비교목록 = textList;
    
    const prompt = `다음 텍스트와 가장 유사한 텍스트를 아래 목록에서 찾아주세요.

찾을 텍스트: "${text1}"

비교할 목록:
${비교목록.map((t, i) => `${i + 1}. ${t}`).join("\n")}

가장 유사한 텍스트의 번호(1부터 시작)와 유사도(0-100)를 JSON 형식으로 반환해주세요:
{
  "index": 번호,
  "similarity": 유사도,
  "reason": "매칭 이유를 간단히 설명"
}

중요: index는 1부터 시작하는 번호입니다.`;

    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini", // 또는 "gpt-3.5-turbo"
      messages: [
        {
          role: "system",
          content: "You are a helpful assistant that finds the most similar text from a list.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
      response_format: { type: "json_object" },
      temperature: 0.3,
    });

    const result = JSON.parse(completion.choices[0].message.content);
    return result;
  } catch (error) {
    console.error("OpenAI API 오류:", error.message);
    return null;
  }
}

/**
 * 텍스트를 임베딩으로 변환 (벡터 유사도 계산용)
 * @param {string} text - 변환할 텍스트
 * @returns {Promise<number[]>} 임베딩 벡터
 */
export async function getEmbedding(text) {
  if (!openai) {
    console.error("❌ OpenAI 클라이언트가 초기화되지 않았습니다. API 키를 확인해주세요.");
    return null;
  }
  
  try {
    const response = await openai.embeddings.create({
      model: "text-embedding-3-small", // 또는 "text-embedding-ada-002"
      input: text,
    });

    return response.data[0].embedding;
  } catch (error) {
    console.error("임베딩 생성 오류:", error.message);
    return null;
  }
}

/**
 * 코사인 유사도 계산
 * @param {number[]} vec1 - 첫 번째 벡터
 * @param {number[]} vec2 - 두 번째 벡터
 * @returns {number} 유사도 (0-1)
 */
export function cosineSimilarity(vec1, vec2) {
  if (!vec1 || !vec2 || vec1.length !== vec2.length) return 0;

  let dotProduct = 0;
  let norm1 = 0;
  let norm2 = 0;

  for (let i = 0; i < vec1.length; i++) {
    dotProduct += vec1[i] * vec2[i];
    norm1 += vec1[i] * vec1[i];
    norm2 += vec2[i] * vec2[i];
  }

  return dotProduct / (Math.sqrt(norm1) * Math.sqrt(norm2));
}

/**
 * 임베딩을 사용하여 가장 유사한 텍스트 찾기
 * @param {string} queryText - 검색할 텍스트
 * @param {string[]} textList - 비교할 텍스트 목록
 * @returns {Promise<Object>} 매칭 결과
 */
export async function findBestMatchWithEmbedding(queryText, textList) {
  try {
    // 쿼리 텍스트 임베딩
    const queryEmbedding = await getEmbedding(queryText);
    if (!queryEmbedding) return null;

    // 각 텍스트의 임베딩과 유사도 계산
    const similarities = [];
    for (let i = 0; i < textList.length; i++) {
      const textEmbedding = await getEmbedding(textList[i]);
      if (textEmbedding) {
        const similarity = cosineSimilarity(queryEmbedding, textEmbedding);
        similarities.push({
          index: i,
          text: textList[i],
          similarity: similarity,
        });
      }
    }

    // 가장 유사한 것 찾기
    const bestMatch = similarities.reduce((best, current) =>
      current.similarity > best.similarity ? current : best
    );

    return bestMatch;
  } catch (error) {
    console.error("임베딩 매칭 오류:", error.message);
    return null;
  }
}

