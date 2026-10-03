import express from 'express';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';

dotenv.config();

export const app = express();

app.use(express.json({ limit: '15mb' }));
app.use(express.urlencoded({ extended: true, limit: '15mb' }));

// Initialize Gemini SDK if API key is provided
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY || 'MOCK_KEY',
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

// Server-side Gemini analyze API
app.post('/api/gemini/analyze', async (req, res) => {
  const { title, content, fact, feeling, finding, future, habits } = req.body;

  if (!content || !fact) {
    return res.status(400).json({ error: 'Mohon isi konten cerita dan fakta kejadian.' });
  }

  try {
    const prompt = `
Analyze the following Indonesian children's story written for the "CERDAS" (Cerita Digital Anak Sempatik) application.
Story Title: "${title || 'Tanpa Judul'}"
Story Content: "${content}"
Reflection 4F:
- Fact (Apa yang terjadi): "${fact}"
- Feeling (Perasaan): "${feeling || ''}"
- Finding (Temuan/Pelajaran): "${finding || ''}"
- Future (Penerapan): "${future || ''}"
Associated 7 Habits of Great Indonesian Kids checked by student: ${JSON.stringify(habits || [])}

Tasks:
1. Provide a warm, extremely positive, child-friendly feedback/encouragement in Indonesian (max 3 sentences) addressing the student as a "Anak Indonesia Hebat" or "Pahlawan Literasi".
2. Generate an automated short summary of the story in Indonesian (max 2 sentences) for teachers and counselors to quickly read.
3. Verify which of the "7 Kebiasaan Anak Indonesia Hebat" are genuinely reflected in the text content.
   The 7 habits are:
   - "Bangun pagi" (Wake up early)
   - "Beribadah" (Worship/Pray)
   - "Berolahraga" (Exercise)
   - "Makan sehat dan bergizi" (Healthy diet)
   - "Gemar belajar" (Love learning)
   - "Bermasyarakat" (Community/Socialize)
   - "Tidur cepat" (Sleep early)
4. Suggest a gorgeous, short, child-friendly illustration prompt in English (e.g. "A beautiful warm watercolor illustration of a cheerful child...") that we can use to generate cover art.
`;

    const response = await ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: prompt,
      config: {
        systemInstruction: 'Anda adalah "Si CERDAS", maskot pendamping literasi digital anak Indonesia yang ceria, penuh energi positif, ramah, dan sangat memotivasi. Anda membantu anak-anak belajar merefleksikan kisah mereka menggunakan 4F (Fact, Feeling, Finding, Future) dan 7 Kebiasaan Baik.',
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            friendlyFeedback: {
              type: Type.STRING,
              description: 'Ulasan positif, ramah, memotivasi anak-anak, dalam Bahasa Indonesia (maksimal 3 kalimat).'
            },
            summary: {
              type: Type.STRING,
              description: 'Ringkasan cerita singkat dalam Bahasa Indonesia (maksimal 2 kalimat).'
            },
            detectedHabits: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
              description: 'Daftar kebiasaan dari 7 kebiasaan yang terbukti/terdeteksi ada di teks cerita.'
            },
            illustrationPrompt: {
              type: Type.STRING,
              description: 'Prompt ilustrasi visual ramah anak dalam Bahasa Inggris (maksimal 1 kalimat).'
            }
          },
          required: ['friendlyFeedback', 'summary', 'detectedHabits', 'illustrationPrompt']
        }
      }
    });

    const responseText = response.text;
    if (!responseText) {
      throw new Error('Gagal menerima tanggapan dari Gemini AI');
    }

    const result = JSON.parse(responseText.trim());
    return res.json(result);
  } catch (error: any) {
    console.warn('Gemini API quota or error, returning friendly fallback analysis:', error?.message);
    const fallbackResult = {
      friendlyFeedback: `Wah, cerita "${title || 'kamu'}" yang ditulis sangat luar biasa hebat! Refleksi 4F-mu sangat matang dan menginspirasi. Teruslah rajin membaca dan menulis ya, Pahlawan Literasi! 🌟✨`,
      summary: `Cerita inspiratif tentang "${title || 'pengalaman siswa'}" yang merefleksikan sikap positif serta pengamalan kebiasaan baik anak Indonesia.`,
      detectedHabits: Array.isArray(habits) && habits.length > 0 ? habits : ["Gemar belajar", "Beribadah"],
      illustrationPrompt: `A vibrant watercolor children book illustration of a cheerful Indonesian student writing a story, bright pastel colors, cute vector style`
    };
    return res.json(fallbackResult);
  }
});

// Server-side Gemini generate image URL
app.post('/api/gemini/generate-illustration', async (req, res) => {
  try {
    const { prompt } = req.body;
    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required.' });
    }

    const finalPrompt = `${prompt}, beautiful cute digital art illustration for children, vibrant cheerful pastel colors, flat vector styled, friendly and warm, high quality, white background`;
    
    try {
      const response = await ai.models.generateContent({
        model: 'gemini-3.1-flash-lite-image',
        contents: {
          parts: [{ text: finalPrompt }]
        },
        config: {
          imageConfig: {
            aspectRatio: '1:1'
          }
        }
      });

      let base64Image = '';
      if (response.candidates?.[0]?.content?.parts) {
        for (const part of response.candidates[0].content.parts) {
          if (part.inlineData) {
            base64Image = `data:image/png;base64,${part.inlineData.data}`;
            break;
          }
        }
      }

      if (base64Image) {
        return res.json({ imageUrl: base64Image });
      } else {
        return res.json({ imageUrl: null, fallbackNeeded: true });
      }
    } catch (innerErr: any) {
      console.warn('Gemini image generation fallback activated (Quota or Limit):', innerErr?.message || innerErr);
      return res.json({ imageUrl: null, fallbackNeeded: true });
    }
  } catch (error: any) {
    console.warn('Illustration outer error, returning fallback:', error?.message);
    return res.json({ imageUrl: null, fallbackNeeded: true });
  }
});

// Server-side Gemini analyze image and generate story
app.post('/api/gemini/analyze-image', async (req, res) => {
  try {
    const { image } = req.body;
    if (!image) {
      return res.status(400).json({ error: 'Data gambar wajib disertakan.' });
    }

    const parts = image.split(',');
    const base64Data = parts[1];
    const mimeType = parts[0].split(':')[1].split(';')[0];

    const response = await ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: [
        {
          inlineData: {
            data: base64Data,
            mimeType: mimeType
          }
        },
        {
          text: `Anda adalah "Si CERDAS", maskot pendamping literasi anak. Lihatlah gambar/foto kegiatan ini dan buatkan sebuah cerita anak inspiratif dan lengkap dengan ulasan refleksi 4F (Fact, Feeling, Finding, Future) dalam Bahasa Indonesia.
Format keluaran wajib JSON objek dengan properti berikut:
- "title": Judul cerita yang menarik, ceria, dan memotivasi anak berdasarkan foto kegiatan tersebut.
- "content": Kisah utama naratif anak yang mengalir indah, ramah anak, dan santun (maksimal 2-3 paragraf).
- "fact": Fakta peristiwa nyata apa yang sebenarnya terjadi pada foto kegiatan tersebut (singkat padat, 1-2 kalimat).
- "feeling": Perasaan bahagia, bersyukur, atau bangga yang dirasakan anak (1-2 kalimat).
- "finding": Pembelajaran atau nilai kebajikan karakter yang ditemukan dari foto/kegiatan tersebut (1-2 kalimat).
- "future": Rencana masa depan atau komitmen nyata untuk mempertahankan kebiasaan baik ini (1-2 kalimat).`
        }
      ],
      config: {
        systemInstruction: 'Anda adalah "Si CERDAS", asisten literasi anak Indonesia yang cerdas, kreatif, positif, dan ramah anak.',
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            title: { type: Type.STRING },
            content: { type: Type.STRING },
            fact: { type: Type.STRING },
            feeling: { type: Type.STRING },
            finding: { type: Type.STRING },
            future: { type: Type.STRING }
          },
          required: ['title', 'content', 'fact', 'feeling', 'finding', 'future']
        }
      }
    });

    const responseText = response.text;
    if (!responseText) {
      throw new Error('Gagal menerima tanggapan dari Gemini AI');
    }

    const result = JSON.parse(responseText.trim());
    return res.json(result);
  } catch (error: any) {
    console.warn('Gemini Image Analysis fallback activated:', error?.message);
    return res.json({
      title: "Kegiatan Hebat Bersama Teman-Teman",
      content: "Hari ini adalah hari yang sangat spesial! Kami berkumpul dan beraktivitas bersama dengan penuh semangat. Banyak hal seru dan bermanfaat yang kami lakukan bersama-sama.",
      fact: "Saya dan teman-teman mengikuti kegiatan bersama dengan penuh antusias.",
      feeling: "Saya merasa sangat senang, bangga, dan bersyukur dapat berpartisipasi.",
      finding: "Saya belajar bahwa kerjasama dan kebersamaan membuat segala hal menjadi lebih ringan dan menyenangkan.",
      future: "Ke depannya, saya akan terus aktif mengikuti kegiatan positif dan saling membantu sesama teman."
    });
  }
});

// Server-side Gemini audio transcription API
app.post('/api/gemini/transcribe', async (req, res) => {
  try {
    const { audio, mimeType } = req.body;
    if (!audio) {
      return res.status(400).json({ error: 'Data audio wajib disertakan.' });
    }

    const base64Data = audio.includes(',') ? audio.split(',')[1] : audio;
    const audioMime = mimeType || (audio.includes(':') ? audio.split(':')[1].split(';')[0] : 'audio/webm');

    console.log(`Transcribing audio with mimeType: ${audioMime}, data length: ${base64Data.length}`);

    const response = await ai.models.generateContent({
      model: 'gemini-3.5-transcribe',
      contents: [
        {
          inlineData: {
            data: base64Data,
            mimeType: audioMime
          }
        },
        {
          text: 'Tolong transkripsikan audio rekaman suara ini ke dalam teks Bahasa Indonesia dengan sangat akurat dan rapi. Hanya kembalikan teks hasil transkripsi saja, tanpa komentar pendahuluan atau penutup apapun.'
        }
      ]
    });

    const responseText = response.text;
    if (!responseText) {
      throw new Error('Gagal menerima transkripsi dari Gemini AI');
    }

    return res.json({ text: responseText.trim() });
  } catch (error: any) {
    console.warn('Gemini Transcription fallback:', error?.message);
    return res.status(500).json({ error: 'Sistem mengalami pembatasan kuota audio sementara. Harap coba beberapa saat lagi atau ketik cerita secara langsung.' });
  }
});
