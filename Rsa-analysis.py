import pandas as pd
from transformers import pipeline
import torch
import re


print("Çok dilli duygu analizi modeli yükleniyor...")
# Daha güçlü ve çok dilli bir model kullanıyoruz.
sentiment_pipeline = pipeline(
    "sentiment-analysis",
    model="cardiffnlp/twitter-xlm-roberta-base-sentiment"
)
print("Model yüklendi.")

try:
    df = pd.read_excel('rsa_analyzed.xlsx')
except FileNotFoundError:
    try:
        df = pd.read_excel('rsa.xlsx')
    except FileNotFoundError:
        print("rsa.xlsx bulunamadı, rsa.xlsx - Sheet1.csv deneniyor...")
        df = pd.read_csv('rsa.xlsx - Sheet1.csv')

# kelime haritası
keyword_map = {
    'merchant quality': [
        'satıcı', 'mağaza', 'firma', 'müşteri hizmetleri', 'ilgi', 'iletişim', 'seller', 'store',
        'customer service', 'communication', 'ilgisiz satıcı', 'hızlı satıcı', 'güvenilir mağaza',
        'soruma cevap vermedi'
    ],
    'product quality': [
        'kalite', 'kaliteli', 'kalitesiz', 'performans', 'kullanışlı', 'işe yarıyor', 'iş görüyor', 'memnun kaldım',
        'tavsiye ederim',
        'beğendim', 'beğenmedim', 'harika', 'mükemmel', 'süper', 'efsane', 'çok iyi', 'çok güzel',
        'bozuk', 'kötü ürün', 'berbat', 'çöp', 'pişmanlık', 'hayal kırıklığı', 'sorunlu', 'çalışmıyor',
        'quality', 'high-quality', 'low-quality', 'performance', 'useful', 'works well', 'satisfied', 'recommend',
        'love it', 'hate it', 'great', 'excellent', 'superb', 'awesome',
        'defective', 'broken', 'bad product', 'terrible', 'garbage', 'regret', 'disappointment', 'doesn\'t work'
        # 'sağlam' kelimesi buradan çıkarıldı çünkü genellikle paketleme ile ilgili kullanılıyor
    ],
    'packaging/delivery quality': [
        'paketleme', 'paket', 'kutu', 'koli', 'ambalaj', 'poşet', 'iyi sarılmış', 'korunaklı',
        'kırık', 'ezik', 'hasarlı', 'dökülmüş', 'akmış', 'yırtık',
        'packaging', 'box', 'package', 'damaged', 'broken', 'leaked', 'well-packaged',
        'sağlam geldi', 'kutu ezikti', 'patlak geldi', 'dökülmüştü'  # <-- Yeni eklendi
    ],
    'delivery time': [
        'hızlı kargo', 'hızlı teslimat', 'zamanında geldi', 'ertesi gün', 'vaktinde',
        'geç geldi', 'yavaş kargo', 'teslim edilmedi',
        'fast delivery', 'shipping', 'arrived on time', 'late delivery', 'slow shipping'
    ],
    'price': [
        'fiyat', 'pahalı', 'ucuz', 'ekonomik', 'indirim', 'f/p', 'fiyat performans', 'ederi', 'parasına değer',
        'price', 'expensive', 'cheap', 'discount', 'value for money'
    ],
    'return policies': ['iade', 'değişim', 'geri gönder', 'return', 'exchange', 'refund'],
    'authenticity (orijinallik)': ['orijinal', 'sahte', 'orjinal değil', 'çakma', 'authentic', 'original', 'fake',
                                   'not original'],
    'smell': ['koku', 'kokusu', 'kokuyor', 'kokmuyor', 'parfüm', 'smell', 'scent', 'fragrance'],
    'son kullanma tarihi': ['skt', 'tarihi geçmiş', 'son kullanma', 'taze', 'bayat', 'expiration date', 'expired',
                            'fresh'],
    'texture': ['doku', 'yapısı', 'kıvam', 'yapışkan', 'yumuşak', 'sert', 'texture', 'consistency', 'sticky', 'soft'],
    'correctness of product (resimle/açıklamayla uyum)': [
        'farklı ürün', 'yanlış ürün', 'resimdeki gibi', 'görseldeki', 'açıklamadaki gibi', 'aynısı geldi',
        'wrong product', 'different product', 'as described', 'as in the picture',
        'farklı geldi', 'yanlış geldi', 'alakasız ürün', 'görselle aynı değil'  # <-- Yeni eklendi
    ],
    'durability (sağlamlık/kalıcılık)': [
        'dayanıklı', 'sağlam', 'kalıcı', 'kalıcılığı', 'uzun ömürlü', 'çabuk bitti', 'hemen bozuldu', 'kop',
        'dayanıksız',
        'durable', 'long-lasting', 'permanence', 'broke quickly', 'wore out'
    ],
    'ease of application': ['kolay sürülüyor', 'zor uygulanıyor', 'pratik', 'kullanışlı', 'easy to apply',
                            'hard to apply', 'practical']
}


#labeling
def map_sentiment(label):
    if label == 'LABEL_2' or label.lower() == 'positive':
        return 'good'
    elif label == 'LABEL_0' or label.lower() == 'negative':
        return 'bad'
    return 'neutral'


# cümleleri ayır
def split_into_sentences(text):
    text = re.sub(r'\n+', '.', text)
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return [s.strip() for s in sentences if s.strip()]


# analiz döngüsü
for index, row in df.iterrows():
    review_text = row['review_text']

    if not isinstance(review_text, str) or not review_text.strip():
        continue

    print(f"\nİncelenen Yorum #{index + 1}: {review_text[:80]}...")

    sentences = split_into_sentences(review_text)

    try:
        overall_sentiment_result = sentiment_pipeline(review_text)[0]
        overall_sentiment_label = map_sentiment(overall_sentiment_result['label'])
        df.at[index, 'overall'] = overall_sentiment_label
    except Exception as e:
        print(f"Genel analizde hata: {e}")
        continue

    found_specific_topic = False

    for sentence in sentences:
        if not sentence: continue

        for column, keywords in keyword_map.items():
            if not keywords: continue

            for keyword in keywords:
                if keyword in sentence.lower():
                    current_value = df.at[index, column]
                    if pd.notna(current_value) and current_value == 'bad':
                        print(f"  -> '{column}' zaten 'bad' olarak işaretli, atlanıyor.")
                        continue

                    sentiment_result = sentiment_pipeline(sentence)[0]
                    label = map_sentiment(sentiment_result['label'])

                    # olumsuzluk kontrolü
                    negation_words = ['değil', 'yok', 'asla', 'beğenmedim', 'olumsuz', 'yaramaz']
                    is_negative_context = any(word in sentence.lower() for word in negation_words)

                    # olumlu - olumsuz
                    if label == 'good' and is_negative_context:
                        label = 'bad'

                    if label == 'neutral' and pd.notna(current_value) and current_value in ['good', 'bad']:
                        continue

                    df.at[index, column] = label
                    print(f"  -> Cümle: '{sentence[:50]}...' | '{keyword}' bulundu. '{column}' -> {label}")
                    found_specific_topic = True
                    # break ifadesi buradan kaldırıldı # <-- DEĞİŞİKLİK

    # kategori
    if not found_specific_topic and pd.isna(df.at[index, 'product quality']):
        # varsayım
        other_topics_keywords = ['kargo', 'teslimat', 'satıcı', 'mağaza', 'iade', 'teslim']
        if not any(keyword in review_text.lower() for keyword in other_topics_keywords):
            df.at[index, 'product quality'] = overall_sentiment_label
            print(f"  -> Spesifik konu yok ve lojistikle ilgili değil. 'product quality' -> {overall_sentiment_label}")

# outputt
df.to_excel('rsa_analyzed_v2_geliştirilmiş.xlsx', index=False)  # <-- Dosya adı karışmaması için değiştirildi

print("\nAnaliz tamamlandı ve sonuçlar 'rsa_analyzed_v2_geliştirilmiş.xlsx' dosyasına kaydedildi.")