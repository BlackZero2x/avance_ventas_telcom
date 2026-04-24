#!/usr/bin/env python3
"""
Script de prueba para diagnosticar envío a Carlos P.
Intenta diferentes formas de envío: por nombre, por ID directo, etc.
"""
import sys
import os
import time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_server"))

from wa_client import WhatsAppClient

wa = WhatsAppClient()

print("=" * 60)
print("TEST 1: Listar contactos que contienen 'Carlos'")
print("=" * 60)
contactos = wa.list_contacts("Carlos")
for c in contactos:
    print(f"  Nombre: {c['name']:<30} ID: {c['id']:<20} Push: {c.get('pushname', 'N/A')}")

print("\n" + "=" * 60)
print("TEST 2: Enviar texto a 'Carlos P.' (por nombre)")
print("=" * 60)
result = wa.send_text("Carlos P.", "🧪 TEST 1: Prueba de texto directo a Carlos P.")
print(f"Resultado: {result}\n")

print("=" * 60)
print("TEST 3: Enviar texto al ID directo (51968035020@c.us)")
print("=" * 60)
time.sleep(3)
result = wa.send_text("51968035020@c.us", "🧪 TEST 2: Prueba de texto directo por ID")
print(f"Resultado: {result}\n")

print("=" * 60)
print("TEST 4: Enviar texto al número sin @c.us (51968035020)")
print("=" * 60)
time.sleep(3)
result = wa.send_text("51968035020", "🧪 TEST 3: Prueba de texto por número")
print(f"Resultado: {result}\n")

print("=" * 60)
print("Revisa tu WhatsApp para ver cuál de los tests llegó")
print("=" * 60)
