#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер e-mail адресов компаний с их официальных сайтов
"""

import re
import sys
import time
import logging
import argparse
from typing import List, Dict, Optional, Tuple
from urllib.parse import urljoin, urlparse
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('parser.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class EmailParser:
    """Класс для парсинга e-mail адресов с сайтов компаний"""
    
    # Приоритетные e-mail адреса (в порядке приоритета)
    PRIORITY_EMAILS = ['info@', 'sales@', 'hello@', 'office@', 'contact@', 'mail@']
    
    # Нежелательные e-mail адреса
    UNWANTED_EMAILS = ['noreply@', 'no-reply@', 'donotreply@', 'support@']
    
    # Регулярное выражение для поиска e-mail
    EMAIL_REGEX = re.compile(
        r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b',
        re.IGNORECASE
    )
    
    def __init__(self, timeout: int = 15, delay: float = 1.0):
        """
        Инициализация парсера
        
        Args:
            timeout: Таймаут запроса в секундах
            delay: Задержка между запросами в секундах
        """
        self.timeout = timeout
        self.delay = delay
        
        # Настройка сессии с повторными попытками
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "HEAD"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)
        
        # Заголовки для имитации браузера
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        })
    
    def normalize_url(self, url: str) -> str:
        """
        Нормализация URL
        
        Args:
            url: Исходный URL
            
        Returns:
            Нормализованный URL
        """
        if not url:
            return ''
        
        # Добавляем протокол, если отсутствует
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        
        return url.rstrip('/')
    
    def get_page_content(self, url: str) -> Optional[BeautifulSoup]:
        """
        Получение содержимого страницы
        
        Args:
            url: URL страницы
            
        Returns:
            BeautifulSoup объект или None при ошибке
        """
        try:
            response = self.session.get(url, timeout=self.timeout, allow_redirects=True)
            response.raise_for_status()
            
            # Проверка кодировки
            if response.encoding is None or response.encoding == 'ISO-8859-1':
                response.encoding = response.apparent_encoding or 'utf-8'
            
            return BeautifulSoup(response.text, 'lxml')
        except requests.exceptions.RequestException as e:
            logger.warning(f"Ошибка при загрузке {url}: {e}")
            return None
        except Exception as e:
            logger.error(f"Неожиданная ошибка при обработке {url}: {e}")
            return None
    
    def extract_emails_from_text(self, text: str) -> List[str]:
        """
        Извлечение e-mail адресов из текста
        
        Args:
            text: Текст для поиска
            
        Returns:
            Список найденных e-mail адресов
        """
        if not text:
            return []
        
        emails = self.EMAIL_REGEX.findall(text)
        # Убираем дубликаты и приводим к нижнему регистру
        unique_emails = list(set(email.lower() for email in emails))
        return unique_emails
    
    def extract_emails_from_soup(self, soup: Optional[BeautifulSoup]) -> List[str]:
        """
        Извлечение e-mail адресов из BeautifulSoup объекта
        Ищет email в HTML коде: mailto ссылки и текст страницы
        
        Args:
            soup: BeautifulSoup объект или None
            
        Returns:
            Список найденных e-mail адресов
        """
        if not soup:
            return []
        
        emails = []
        
        # 1. Поиск в mailto ссылках (приоритетный способ)
        for link in soup.find_all('a', href=True):
            href = link.get('href', '').strip()
            if href.startswith('mailto:'):
                email = href.replace('mailto:', '').split('?')[0].strip()
                if email:
                    emails.append(email.lower())
        
        # 2. Поиск в тексте страницы (regex поиск)
        page_text = soup.get_text()
        emails.extend(self.extract_emails_from_text(page_text))
        
        # Убираем дубликаты
        return list(set(emails))
    
    def filter_emails(self, emails: List[str]) -> List[str]:
        """
        Фильтрация e-mail адресов
        
        Args:
            emails: Список e-mail адресов
            
        Returns:
            Отфильтрованный список
        """
        if not emails:
            return []
        
        # Убираем нежелательные
        filtered = [
            email for email in emails
            if not any(unwanted in email.lower() for unwanted in self.UNWANTED_EMAILS)
        ]
        
        # Если после фильтрации ничего не осталось, возвращаем исходный список
        if not filtered:
            filtered = emails
        
        return filtered
    
    def select_primary_email(self, emails: List[str]) -> Optional[str]:
        """
        Выбор основного e-mail адреса из списка
        
        Args:
            emails: Список e-mail адресов
            
        Returns:
            Основной e-mail или None
        """
        if not emails:
            return None
        
        # Ищем по приоритету
        for priority in self.PRIORITY_EMAILS:
            for email in emails:
                if email.startswith(priority):
                    return email
        
        # Если не нашли приоритетный, возвращаем первый
        return emails[0]
    
    def find_company_email(self, website: str) -> List[str]:
        """
        Поиск e-mail адресов компании на её сайте
        Просто ищет email в HTML коде страницы
        
        Args:
            website: URL сайта компании
            
        Returns:
            Список найденных e-mail адресов (может быть пустым)
        """
        website = self.normalize_url(website)
        if not website:
            return []
        
        # Просто открываем страницу и ищем email в HTML
        soup = self.get_page_content(website)
        if not soup:
            return []
        
        # Извлекаем все email из HTML кода страницы
        all_emails = self.extract_emails_from_soup(soup)
        
        # Убираем дубликаты
        unique_emails = list(set(all_emails))
        
        # Фильтруем нежелательные
        filtered_emails = self.filter_emails(unique_emails)
        
        # Сортируем по приоритету (приоритетные первыми)
        sorted_emails = sorted(filtered_emails, key=lambda email: (
            min([i for i, priority in enumerate(self.PRIORITY_EMAILS) if email.startswith(priority)], default=999),
            email
        ))
        
        if sorted_emails:
            logger.info(f"Найдено {len(sorted_emails)} email для {website}: {', '.join(sorted_emails)}")
        else:
            logger.warning(f"Email не найден для {website}")
        
        return sorted_emails
    
    def extract_company_name_from_page(self, page_url: str) -> str:
        """
        Извлечение названия компании со страницы
        
        Args:
            page_url: URL страницы компании
            
        Returns:
            Название компании или пустая строка
        """
        soup = self.get_page_content(page_url)
        if not soup:
            return ''
        
        # Пробуем найти название в различных местах
        # 1. В заголовке h1
        h1 = soup.find('h1')
        if h1:
            name = h1.get_text(strip=True)
            if name and len(name) < 200:  # Разумная длина для названия
                return name
        
        # 2. В title
        title = soup.find('title')
        if title:
            name = title.get_text(strip=True)
            if name and len(name) < 200:
                return name
        
        # 3. В мета-тегах
        meta_title = soup.find('meta', property='og:title')
        if meta_title and meta_title.get('content'):
            name = meta_title.get('content').strip()
            if name and len(name) < 200:
                return name
        
        # 4. В первом значимом заголовке
        for tag in ['h2', 'h3', 'h4']:
            header = soup.find(tag)
            if header:
                name = header.get_text(strip=True)
                if name and len(name) < 200:
                    return name
        
        return ''
    
    def extract_company_data_from_page(self, page_url: str, aggregator_domain: str) -> Tuple[List[Dict[str, List[str]]], List[str]]:
        """
        Извлечение данных о компаниях со страницы агрегатора
        
        Args:
            page_url: URL страницы агрегатора
            aggregator_domain: Домен агрегатора (для фильтрации)
            
        Returns:
            Кортеж: (список словарей с данными компаний, список внутренних ссылок для обхода)
            Формат: ([{'company_url': url, 'emails': [emails], 'company_name': name}], [internal_links])
        """
        soup = self.get_page_content(page_url)
        if not soup:
            return [], []
        
        # Убираем www для сравнения
        aggregator_domain_clean = aggregator_domain[4:] if aggregator_domain.startswith('www.') else aggregator_domain
        
        company_data = {}  # {company_url: [emails]}
        internal_links = []  # Внутренние ссылки для дальнейшего обхода
        
        # Ищем все ссылки на странице (включая относительные)
        for link in soup.find_all('a', href=True):
            href = link.get('href', '').strip()
            if not href:
                continue
            
            # Пропускаем mailto, javascript, якоря, tel
            if href.startswith(('mailto:', 'javascript:', '#', 'tel:')):
                continue
            
            # Преобразуем относительные ссылки в абсолютные
            full_url = urljoin(page_url, href)
            
            try:
                parsed = urlparse(full_url)
                link_domain = parsed.netloc.lower()
                
                # Пропускаем ссылки без домена (не должно быть, но на всякий случай)
                if not link_domain:
                    continue
                
                # Пропускаем не-HTTP(S) ссылки
                if parsed.scheme not in ('http', 'https'):
                    continue
                
                # Убираем www для сравнения
                link_domain_clean = link_domain[4:] if link_domain.startswith('www.') else link_domain
                
                # Внутренняя ссылка агрегатора
                if link_domain_clean == aggregator_domain_clean:
                    # Проверяем, похожа ли ссылка на страницу компании (например, /partners/...)
                    path = parsed.path.lower()
                    # Если это ссылка на страницу партнёра/компании, добавляем в список для обработки
                    if any(pattern in path for pattern in ['/partner', '/company', '/intro', '/detail']):
                        if full_url not in internal_links and full_url != page_url:
                            internal_links.append(full_url)
                    elif full_url not in internal_links and full_url != page_url:
                        # Обычная внутренняя ссылка для обхода
                        internal_links.append(full_url)
                else:
                    # Внешняя ссылка - это ссылка на компанию
                    if full_url not in company_data:
                        company_data[full_url] = []
            except Exception as e:
                logger.debug(f"Ошибка при обработке ссылки {href}: {e}")
                continue
        
        # Ищем email адреса прямо на этой странице
        page_emails = self.extract_emails_from_soup(soup)
        if page_emails:
            filtered_emails = self.filter_emails(page_emails)
            # Если на странице есть email и есть ссылки на компании, привязываем email к компаниям
            if filtered_emails and company_data:
                for company_url in company_data:
                    company_data[company_url].extend(filtered_emails)
            # Если есть email, но нет внешних ссылок - это может быть страница с информацией о компании
            # но без прямой ссылки на её сайт. В этом случае email остаётся без привязки к конкретной компании.
        
        # Формируем результат - только компании с URL
        result = []
        for company_url, emails in company_data.items():
            if company_url:  # Только если есть URL компании
                unique_emails = list(set(emails)) if emails else []
                result.append({
                    'company_url': company_url, 
                    'emails': unique_emails,
                    'company_name': ''  # Будет заполнено позже при обработке страницы
                })
        
        return result, internal_links
    
    def parse_with_depth(self, website_url: str, depth: int = 0) -> List[Dict[str, str]]:
        """
        Универсальная функция парсинга с параметром глубины
        
        Args:
            website_url: URL сайта для парсинга
            depth: Глубина поиска (0 - сразу собираем почты, 1+ - собираем ссылки и переходим по ним)
            
        Returns:
            Список словарей с результатами
        """
        website_url = self.normalize_url(website_url)
        if not website_url:
            return []
        
        results = []
        
        if depth == 0:
            # Глубина 0: сразу собираем почты с сайта
            logger.info(f"Глубина 0: сбор email с сайта {website_url}")
            emails = self.find_company_email(website_url)
            company_name = self.extract_company_name_from_page(website_url)
            
            email_str = ', '.join(emails) if emails else '-'
            results.append({
                'Company Name': company_name or '-',
                'Website': website_url,
                'Email': email_str
            })
        else:
            # Глубина 1+: собираем ссылки и переходим по ним
            logger.info(f"Глубина {depth}: сбор ссылок с сайта {website_url}")
            company_data_list = self.extract_company_links_from_aggregator(website_url, max_depth=depth)
            
            for company_data in company_data_list:
                company_url = company_data['company_url']
                emails_from_page = company_data.get('emails', [])
                company_name = company_data.get('company_name', '')
                
                # Парсим email с сайта компании
                emails_from_site = []
                if company_url:
                    parsed_url = urlparse(company_url)
                    url_domain = parsed_url.netloc.lower()
                    url_domain_clean = url_domain[4:] if url_domain.startswith('www.') else url_domain
                    site_domain_clean = urlparse(website_url).netloc.lower()
                    site_domain_clean = site_domain_clean[4:] if site_domain_clean.startswith('www.') else site_domain_clean
                    
                    # Если это внешний сайт, парсим email с него
                    if url_domain_clean != site_domain_clean:
                        emails_from_site = self.find_company_email(company_url)
                        if not company_name:
                            company_name = self.extract_company_name_from_page(company_url)
                
                # Объединяем все email
                all_emails = list(set(emails_from_page + emails_from_site))
                sorted_emails = sorted(all_emails, key=lambda email: (
                    min([i for i, priority in enumerate(self.PRIORITY_EMAILS) if email.startswith(priority)], default=999),
                    email
                ))
                
                email_str = ', '.join(sorted_emails) if sorted_emails else '-'
                results.append({
                    'Company Name': company_name or '-',
                    'Website': company_url,
                    'Email': email_str
                })
                
                time.sleep(self.delay)
        
        return results
    
    def parse_sites_only(self, website_url: str, depth: int = 0) -> List[Dict[str, str]]:
        """
        Парсинг только сайтов с названиями (без парсинга email)
        
        Args:
            website_url: URL сайта для парсинга
            depth: Глубина поиска (0 - один сайт, 1+ - собираем ссылки и переходим по ним)
            
        Returns:
            Список словарей с результатами (только названия и сайты)
        """
        website_url = self.normalize_url(website_url)
        if not website_url:
            return []
        
        results = []
        
        if depth == 0:
            # Глубина 0: просто название и сайт
            logger.info(f"Глубина 0: сбор информации о сайте {website_url}")
            company_name = self.extract_company_name_from_page(website_url)
            
            results.append({
                'Company Name': company_name or '-',
                'Website': website_url
            })
        else:
            # Глубина 1+: собираем ссылки и переходим по ним
            logger.info(f"Глубина {depth}: сбор ссылок с сайта {website_url}")
            company_data_list = self.extract_company_links_from_aggregator(website_url, max_depth=depth)
            
            for company_data in company_data_list:
                company_url = company_data['company_url']
                company_name = company_data.get('company_name', '')
                
                # Если название не найдено, пытаемся извлечь
                if not company_name and company_url:
                    parsed_url = urlparse(company_url)
                    url_domain = parsed_url.netloc.lower()
                    url_domain_clean = url_domain[4:] if url_domain.startswith('www.') else url_domain
                    site_domain_clean = urlparse(website_url).netloc.lower()
                    site_domain_clean = site_domain_clean[4:] if site_domain_clean.startswith('www.') else site_domain_clean
                    
                    # Если это внешний сайт, пытаемся извлечь название
                    if url_domain_clean != site_domain_clean:
                        company_name = self.extract_company_name_from_page(company_url)
                
                results.append({
                    'Company Name': company_name or '-',
                    'Website': company_url
                })
                
                time.sleep(self.delay)
        
        return results
    
    def extract_company_links_from_aggregator(self, aggregator_url: str, max_depth: int = 3) -> List[Dict[str, any]]:
        """
        Извлечение ссылок на компании со страницы агрегатора с обходом внутренних страниц
        
        Args:
            aggregator_url: URL страницы агрегатора
            max_depth: Максимальная глубина обхода внутренних страниц (по умолчанию 3)
            
        Returns:
            Список словарей: [{'company_url': url, 'emails': [emails], 'company_name': name}]
        """
        aggregator_url = self.normalize_url(aggregator_url)
        if not aggregator_url:
            return []
        
        # Получаем домен агрегатора
        aggregator_domain = urlparse(aggregator_url).netloc.lower()
        
        all_company_data = {}  # {company_url: {'emails': set(), 'company_name': str}}
        visited_pages = set()  # Посещённые страницы
        pages_to_visit = [(aggregator_url, 0)]  # (url, depth)
        
        logger.info(f"Начинаем обход страниц агрегатора {aggregator_url} (максимальная глубина: {max_depth})")
        
        while pages_to_visit:
            current_url, depth = pages_to_visit.pop(0)
            
            # Пропускаем, если уже посещали или превышена глубина
            if current_url in visited_pages or depth > max_depth:
                continue
            
            visited_pages.add(current_url)
            logger.info(f"Обработка страницы (глубина {depth}): {current_url}")
            
            # Извлекаем данные со страницы
            company_data_list, internal_links = self.extract_company_data_from_page(current_url, aggregator_domain)
            
            # Обрабатываем найденные компании
            for item in company_data_list:
                company_url = item['company_url']
                emails = item.get('emails', [])
                
                if company_url not in all_company_data:
                    all_company_data[company_url] = {'emails': set(), 'company_name': ''}
                
                all_company_data[company_url]['emails'].update(emails)
            
            # Проверяем, является ли текущая страница страницей компании (например, /partners/...)
            # Если да, извлекаем название и email с неё
            parsed_current = urlparse(current_url)
            path = parsed_current.path.lower()
            if any(pattern in path for pattern in ['/partner', '/company', '/intro', '/detail']):
                # Это страница компании на агрегаторе
                soup_page = self.get_page_content(current_url)
                company_emails = []
                if soup_page:
                    company_emails = self.extract_emails_from_soup(soup_page)
                if company_emails:
                    filtered_emails = self.filter_emails(company_emails)
                    company_name = self.extract_company_name_from_page(current_url)
                    
                    # Ищем внешнюю ссылку на сайт компании на этой странице
                    soup = self.get_page_content(current_url)
                    if soup:
                        # Ищем ссылки на внешние сайты
                        for link in soup.find_all('a', href=True):
                            href = link.get('href', '').strip()
                            if href.startswith(('http://', 'https://')):
                                parsed_link = urlparse(href)
                                link_domain = parsed_link.netloc.lower()
                                link_domain_clean = link_domain[4:] if link_domain.startswith('www.') else link_domain
                                aggregator_domain_clean = aggregator_domain[4:] if aggregator_domain.startswith('www.') else aggregator_domain
                                
                                if link_domain_clean != aggregator_domain_clean:
                                    # Нашли внешнюю ссылку на сайт компании
                                    if href not in all_company_data:
                                        all_company_data[href] = {'emails': set(), 'company_name': company_name}
                                    all_company_data[href]['emails'].update(filtered_emails)
                                    if not all_company_data[href]['company_name']:
                                        all_company_data[href]['company_name'] = company_name
                                    break
                        else:
                            # Если не нашли внешнюю ссылку, создаём запись с URL страницы агрегатора
                            if current_url not in all_company_data:
                                all_company_data[current_url] = {'emails': set(), 'company_name': company_name}
                            all_company_data[current_url]['emails'].update(filtered_emails)
            
            # Добавляем внутренние ссылки для обхода
            if depth < max_depth:
                for internal_link in internal_links:
                    if internal_link not in visited_pages:
                        pages_to_visit.append((internal_link, depth + 1))
            
            # Пауза между запросами
            if pages_to_visit:
                time.sleep(self.delay)
        
        # Преобразуем в список словарей
        result = []
        for company_url, data in all_company_data.items():
            emails_list = sorted(list(data['emails']))
            result.append({
                'company_url': company_url, 
                'emails': emails_list,
                'company_name': data.get('company_name', '')
            })
        
        logger.info(f"Найдено {len(result)} уникальных компаний на агрегаторе {aggregator_url}")
        return result
    
    def parse_companies(self, companies: List[Dict[str, str]]) -> List[Dict[str, str]]:
        """
        Парсинг e-mail адресов для списка компаний
        
        Args:
            companies: Список словарей с ключами 'company_name' и 'company_website'
            
        Returns:
            Список словарей с результатами
        """
        results = []
        total = len(companies)
        
        for idx, company in enumerate(companies, 1):
            company_name = company.get('company_name', '')
            company_website = company.get('company_website', '')
            
            logger.info(f"[{idx}/{total}] Обработка: {company_name} ({company_website})")
            
            emails = []
            if company_website:
                emails = self.find_company_email(company_website)
                # Пауза между запросами
                if idx < total:
                    time.sleep(self.delay)
            
            # Форматируем email: если найдены - через запятую, если нет - "-"
            email_str = ', '.join(emails) if emails else '-'
            
            results.append({
                'Company Name': company_name,
                'Website': company_website,
                'Email': email_str
            })
        
        return results
    
    def parse_aggregators(self, aggregators: List[Dict[str, str]]) -> List[Dict[str, str]]:
        """
        Парсинг компаний со страниц агрегаторов и их e-mail адресов
        
        Args:
            aggregators: Список словарей с ключами 'company_name' (название агрегатора) 
                        и 'company_website' (ссылка на агрегатор)
            
        Returns:
            Список словарей с результатами
        """
        results = []
        total_aggregators = len(aggregators)
        
        for agg_idx, aggregator in enumerate(aggregators, 1):
            agg_name = aggregator.get('company_name', '')
            agg_url = aggregator.get('company_website', '')
            
            logger.info(f"[Агрегатор {agg_idx}/{total_aggregators}] Обработка: {agg_name} ({agg_url})")
            
            # Извлекаем данные о компаниях со страниц агрегатора (с обходом внутренних страниц)
            company_data_list = self.extract_company_links_from_aggregator(agg_url)
            
            if not company_data_list:
                logger.warning(f"Не найдено компаний на агрегаторе {agg_url}")
                # Добавляем запись без компаний
                results.append({
                    'Aggregator Name': agg_name,
                    'Aggregator URL': agg_url,
                    'Company Name': '-',
                    'Company Website': '',
                    'Email': '-'
                })
                continue
            
            # Обрабатываем каждую найденную компанию
            total_companies = len(company_data_list)
            for comp_idx, company_data in enumerate(company_data_list, 1):
                company_url = company_data['company_url']
                emails_from_aggregator = company_data.get('emails', [])
                company_name = company_data.get('company_name', '')
                
                logger.info(f"  [{comp_idx}/{total_companies}] Обработка компании: {company_name or company_url}")
                
                # Парсим email с сайта компании (только если это внешний сайт, не страница агрегатора)
                emails_from_company_site = []
                parsed_url = urlparse(company_url)
                url_domain = parsed_url.netloc.lower()
                url_domain_clean = url_domain[4:] if url_domain.startswith('www.') else url_domain
                agg_domain_clean = urlparse(agg_url).netloc.lower()
                agg_domain_clean = agg_domain_clean[4:] if agg_domain_clean.startswith('www.') else agg_domain_clean
                
                # Если это внешний сайт компании, парсим email с него
                if company_url and url_domain_clean != agg_domain_clean:
                    emails_from_company_site = self.find_company_email(company_url)
                    # Если не нашли название на агрегаторе, пытаемся извлечь с сайта компании
                    if not company_name:
                        company_name = self.extract_company_name_from_page(company_url)
                
                # Объединяем email с агрегатора и с сайта компании
                all_emails = list(set(emails_from_aggregator + emails_from_company_site))
                
                # Сортируем по приоритету
                sorted_emails = sorted(all_emails, key=lambda email: (
                    min([i for i, priority in enumerate(self.PRIORITY_EMAILS) if email.startswith(priority)], default=999),
                    email
                ))
                
                # Форматируем email: если найдены - через запятую, если нет - "-"
                email_str = ', '.join(sorted_emails) if sorted_emails else '-'
                
                results.append({
                    'Aggregator Name': agg_name,
                    'Aggregator URL': agg_url,
                    'Company Name': company_name or '-',
                    'Company Website': company_url,
                    'Email': email_str
                })
                
                # Пауза между запросами
                if comp_idx < total_companies or agg_idx < total_aggregators:
                    time.sleep(self.delay)
        
        return results


def load_input_data(file_path: str) -> List[Dict[str, str]]:
    """
    Загрузка входных данных из файла
    
    Args:
        file_path: Путь к файлу (Excel, CSV или TXT)
        
    Returns:
        Список словарей с данными компаний
    """
    file_path = Path(file_path)
    
    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    
    extension = file_path.suffix.lower()
    
    try:
        if extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        elif extension == '.csv':
            df = pd.read_csv(file_path, encoding='utf-8')
        elif extension == '.txt':
            # Простой формат: название|сайт или название,сайт
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    parts = line.replace('|', ',').split(',')
                    if len(parts) >= 2:
                        data.append({
                            'company_name': parts[0].strip(),
                            'company_website': parts[1].strip()
                        })
            df = pd.DataFrame(data)
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {extension}")
        
        # Нормализация названий колонок
        df.columns = df.columns.str.lower().str.strip()
        
        # Проверка наличия необходимых колонок
        required_columns = ['company_name', 'company_website']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Отсутствуют необходимые колонки: {missing_columns}")
        
        # Преобразование в список словарей
        companies = df[required_columns].to_dict('records')
        
        logger.info(f"Загружено {len(companies)} компаний из {file_path}")
        return companies
    
    except Exception as e:
        logger.error(f"Ошибка при загрузке файла {file_path}: {e}")
        raise


def save_results(results: List[Dict[str, str]], output_path: str, mode: str = 'companies', sites_only: bool = False):
    """
    Сохранение результатов в Excel файл
    
    Args:
        results: Список словарей с результатами
        output_path: Путь к выходному файлу
        mode: Режим работы ('companies' или 'aggregators')
        sites_only: Если True, сохраняет только названия и сайты (без email)
    """
    df = pd.DataFrame(results)
    
    # Убеждаемся, что колонки в правильном порядке
    if sites_only:
        # Режим: только сайты с названиями
        if mode == 'aggregators':
            columns = ['Aggregator Name', 'Aggregator URL', 'Company Name', 'Company Website']
        else:
            columns = ['Company Name', 'Website']
    else:
        # Обычный режим: с email
        if mode == 'aggregators':
            columns = ['Aggregator Name', 'Aggregator URL', 'Company Name', 'Company Website', 'Email']
        else:
            columns = ['Company Name', 'Website', 'Email']
    
    # Проверяем наличие всех колонок
    available_columns = [col for col in columns if col in df.columns]
    df = df[available_columns]
    
    # Сохранение в Excel
    output_path = Path(output_path)
    df.to_excel(output_path, index=False, engine='openpyxl')
    
    # Статистика
    total = len(results)
    
    logger.info(f"Результаты сохранены в {output_path}")
    if sites_only:
        logger.info(f"Всего записей: {total} (только сайты с названиями)")
    else:
        # Считаем записи с email (не "-" и не пустые)
        with_email = sum(1 for r in results if r.get('Email') and r.get('Email') != '-')
        if mode == 'aggregators':
            unique_companies = len(set(r.get('Company Website', '') for r in results if r.get('Company Website')))
            logger.info(f"Всего записей: {total}, уникальных компаний: {unique_companies}, с email: {with_email}, без email: {total - with_email}")
        else:
            logger.info(f"Всего компаний: {total}, с email: {with_email}, без email: {total - with_email}")


def get_files_in_directory(directory: str = '.') -> List[str]:
    """
    Получение списка файлов в директории
    
    Args:
        directory: Путь к директории
        
    Returns:
        Список имён файлов
    """
    try:
        path = Path(directory)
        files = [f.name for f in path.iterdir() if f.is_file()]
        return sorted(files)
    except Exception as e:
        logger.warning(f"Ошибка при чтении директории: {e}")
        return []


def select_file_by_number(files: List[str]) -> Optional[str]:
    """
    Выбор файла по номеру из списка
    
    Args:
        files: Список имён файлов
        
    Returns:
        Выбранное имя файла или None
    """
    if not files:
        return None
    
    while True:
        choice = input(f"Введите номер файла (1-{len(files)}) или название файла: ").strip()
        
        # Проверяем, является ли ввод числом
        if choice.isdigit():
            file_num = int(choice)
            if 1 <= file_num <= len(files):
                selected_file = files[file_num - 1]
                print(f"Выбран файл: {selected_file}\n")
                return selected_file
            else:
                print(f"Ошибка: номер должен быть от 1 до {len(files)}. Попробуйте снова.")
                continue
        
        # Если не число, проверяем как название файла
        if choice:
            if choice in files:
                print(f"Выбран файл: {choice}\n")
                return choice
            else:
                # Проверяем существование файла
                if Path(choice).exists():
                    return choice
                else:
                    print(f"Ошибка: файл '{choice}' не найден. Попробуйте снова.")
                    continue
        else:
            print("Ошибка: введите номер или название файла. Попробуйте снова.")


def interactive_input() -> Tuple[str, str, int, float, bool]:
    """
    Интерактивный ввод параметров
    
    Returns:
        Кортеж (site_url, output_file, depth, timeout, delay, sites_only)
    """
    print("\n" + "="*60)
    print("Парсер e-mail адресов компаний")
    print("="*60 + "\n")
    
    # Выбор режима работы
    print("Выберите режим работы:")
    print("  1. Парсинг сайтов (только сайты с названиями, без email)")
    print("  2. Парсинг почт (сайты с названиями и email адресами)")
    print()
    
    while True:
        mode_choice = input("Введите номер режима (1 или 2): ").strip()
        if mode_choice == '1':
            sites_only = True
            mode_name = "Парсинг сайтов"
            break
        elif mode_choice == '2':
            sites_only = False
            mode_name = "Парсинг почт"
            break
        else:
            print("Ошибка: введите 1 или 2. Попробуйте снова.")
    
    print(f"\nВыбран режим: {mode_name}\n")
    
    # Запрос ссылки на сайт
    while True:
        site_url = input("Введите ссылку на сайт: ").strip()
        if not site_url:
            print("Ошибка: ссылка на сайт не может быть пустой. Попробуйте снова.")
            continue
        break
    
    # Запрос глубины поиска
    while True:
        depth_input = input("Глубина поиска (0 - сразу собираем данные, 1+ - собираем ссылки и переходим по ним): ").strip()
        try:
            depth = int(depth_input) if depth_input else 0
            if depth < 0:
                print("Ошибка: глубина должна быть неотрицательным числом. Попробуйте снова.")
                continue
            break
        except ValueError:
            print("Ошибка: введите число. Попробуйте снова.")
            continue
    
    # Показываем файлы в текущей директории
    files = get_files_in_directory()
    if files:
        print("\nДоступные файлы в текущей директории:")
        for i, file in enumerate(files, 1):
            print(f"  {i}. {file}")
        print()
    
    # Запрос выходного файла
    while True:
        if files:
            output_file = select_file_by_number(files)
            if output_file:
                # Если выбран существующий файл, спрашиваем подтверждение
                confirm = input(f"Использовать '{output_file}' как выходной файл? (y/n): ").strip().lower()
                if confirm == 'y':
                    break
                else:
                    output_file = input("Введите название выходного файла (Excel): ").strip()
            else:
                output_file = input("Введите название выходного файла (Excel): ").strip()
        else:
            output_file = input("Введите название выходного файла (Excel): ").strip()
        
        if not output_file:
            print("Ошибка: имя файла не может быть пустым. Попробуйте снова.")
            continue
        
        # Добавляем расширение .xlsx, если не указано
        if not output_file.endswith(('.xlsx', '.xls')):
            output_file += '.xlsx'
        
        break
    
    # Запрос таймаута (опционально)
    timeout_input = input("Таймаут запроса в секундах (Enter для значения по умолчанию 15): ").strip()
    timeout = int(timeout_input) if timeout_input.isdigit() else 15
    
    # Запрос задержки (опционально)
    delay_input = input("Задержка между запросами в секундах (Enter для значения по умолчанию 1.0): ").strip()
    try:
        delay = float(delay_input) if delay_input else 1.0
    except ValueError:
        delay = 1.0
    
    return site_url, output_file, depth, timeout, delay, sites_only


def main():
    """Главная функция для запуска из командной строки"""
    parser = argparse.ArgumentParser(
        description='Парсер e-mail адресов компаний с их официальных сайтов'
    )
    parser.add_argument(
        'site_url',
        nargs='?',
        help='URL сайта для парсинга. Если не указан, будет интерактивный режим.'
    )
    parser.add_argument(
        'output_file',
        nargs='?',
        help='Путь к выходному Excel файлу. Если не указан, будет интерактивный режим.'
    )
    parser.add_argument(
        '--sites-only',
        action='store_true',
        help='Режим: только сайты с названиями (без парсинга email)'
    )
    parser.add_argument(
        '--depth',
        type=int,
        default=0,
        help='Глубина поиска (0 - сразу собираем данные, 1+ - собираем ссылки и переходим по ним)'
    )
    parser.add_argument(
        '--timeout',
        type=int,
        default=15,
        help='Таймаут запроса в секундах (по умолчанию: 15)'
    )
    parser.add_argument(
        '--delay',
        type=float,
        default=1.0,
        help='Задержка между запросами в секундах (по умолчанию: 1.0)'
    )
    
    args = parser.parse_args()
    
    # Если аргументы не переданы, используем интерактивный режим
    if not args.site_url or not args.output_file:
        site_url, output_file, depth, timeout, delay, sites_only = interactive_input()
    else:
        site_url = args.site_url
        output_file = args.output_file
        depth = args.depth
        timeout = args.timeout
        delay = args.delay
        sites_only = args.sites_only
    
    try:
        # Создание парсера
        email_parser = EmailParser(timeout=timeout, delay=delay)
        
        # Парсинг в зависимости от режима
        if sites_only:
            logger.info("Режим: Парсинг сайтов (только названия и сайты)")
            results = email_parser.parse_sites_only(site_url, depth=depth)
        else:
            logger.info("Режим: Парсинг почт (сайты с названиями и email)")
            results = email_parser.parse_with_depth(site_url, depth=depth)
        
        if not results:
            logger.warning("Результаты не найдены")
            return
        
        # Сохранение результатов
        save_results(results, output_file, mode='companies', sites_only=sites_only)
        
        logger.info("Парсинг завершён успешно!")
    
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
