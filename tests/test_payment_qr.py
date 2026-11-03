"""Tests for ground truth taken from payment QR codes and archive file names.

Every accuracy number measured against these labels depends on the decoder,
so it is tested against payloads built here with the published PAY by square
layout (fake data only).
"""

import lzma
import zlib

import pytest

from evaluation import payment_qr as qr

ALPHABET = "0123456789ABCDEFGHIJKLMNOPQRSTUV"
FILTERS = [{"id": lzma.FILTER_LZMA1, "lc": 3, "lp": 0, "pb": 2, "dict_size": 128 * 1024}]
FAKE_IBAN = "SK0000000000000000000000"


def encode_pay_by_square(fields: list[str], corrupt_crc: bool = False) -> str:
    payload = "\t".join(fields).encode()
    crc = zlib.crc32(payload) ^ (1 if corrupt_crc else 0)
    data = crc.to_bytes(4, "little") + payload
    raw = b"\x00\x00" + len(data).to_bytes(2, "little") + lzma.compress(data, format=lzma.FORMAT_RAW, filters=FILTERS)
    bits = "".join(f"{b:08b}" for b in raw)
    bits += "0" * (-len(bits) % 5)
    return "".join(ALPHABET[int(bits[i:i + 5], 2)] for i in range(0, len(bits), 5))


def pay_fields(amount="302.58", due="20260808", vs="20260384"):
    return ["", "1", "1", amount, "EUR", due, vs, "0308", "", "", "", "1", FAKE_IBAN, "BICCODEX", "0", "0"]


class TestPayBySquare:
    def test_round_trip(self):
        decoded = qr.decode_pay_by_square(encode_pay_by_square(pay_fields()))
        assert decoded == {"amount": "302.58", "currency": "EUR", "due_date": "08.08.2026",
                           "variable_symbol": "20260384", "iban": FAKE_IBAN}

    def test_missing_due_date_stays_empty(self):
        assert qr.decode_pay_by_square(encode_pay_by_square(pay_fields(due="")))["due_date"] == ""

    def test_checksum_mismatch_is_rejected(self):
        assert qr.decode_pay_by_square(encode_pay_by_square(pay_fields(), corrupt_crc=True)) is None

    @pytest.mark.parametrize("text", ["https://example.com/pay", "", "ABC"])
    def test_other_qr_content_is_not_a_payment(self, text):
        assert qr.decode_payment_qr(text) is None


class TestQrPlatba:
    def test_czech_spd_code(self):
        text = "SPD*1.0*ACC:CZ0000000000000000000000*AM:0000962.70*CC:EUR*DT:20260625*X-VS:4402600155*MSG:A%2AB"
        decoded = qr.decode_payment_qr(text)
        assert (decoded["amount"], decoded["due_date"], decoded["variable_symbol"]) == (
            "962.70", "25.06.2026", "4402600155")

    def test_spd_without_amount_is_ignored(self):
        assert qr.decode_spd("SPD*1.0*ACC:CZ0000000000000000000000*X-VS:1") is None


class TestFileNames:
    @pytest.mark.parametrize("name", ["Faktúra 20260384.pdf", "Faktura_20260403.pdf", "Faktúra  20260405.pdf"])
    def test_issued_invoice_number_from_name(self, name):
        labels = qr.labels_from_name(name)
        assert labels["typ"] == "Vystavená" and labels["cislo_faktury"] == name[-12:-4]

    @pytest.mark.parametrize("name", ["DF0451-26.pdf", "zDF0084-26.pdf", "DF 1429 - X.pdf"])
    def test_internal_number_means_received(self, name):
        assert qr.labels_from_name(name) == {"typ": "Prijatá"}

    @pytest.mark.parametrize("name", ["Faktúra 2026070627.pdf", "1..pdf", "scan_012275.pdf"])
    def test_other_names_give_no_labels(self, name):
        # a supplier's 10-digit number must not be mistaken for an issued invoice
        assert qr.labels_from_name(name) == {}


class TestCombiningSources:
    def payment(self, amount="10.00", vs="1"):
        return {"amount": amount, "variable_symbol": vs, "due_date": "01.07.2026", "currency": "EUR", "iban": ""}

    def test_repeated_identical_code_is_fine(self):
        labels, note = qr.labels_from_payments([self.payment(), self.payment()])
        assert labels["suma_s_dph"] == "10.00" and note == "payment QR" and "iban" in labels

    def test_two_different_codes_give_no_labels(self):
        labels, note = qr.labels_from_payments([self.payment(), self.payment(amount="99.00")])
        assert labels == {} and note.startswith("ambiguous")

    def test_issued_name_and_qr_must_agree(self, monkeypatch, tmp_path):
        ok, clash = tmp_path / "Faktúra 20260384.pdf", tmp_path / "Faktúra 20260385.pdf"
        codes = {ok.name: [self.payment(vs="20260384")], clash.name: [self.payment(vs="99999999")]}
        monkeypatch.setattr(qr, "read_payment_codes", lambda pdf: codes[pdf.name])

        rows, stats = qr.build_labels([ok, clash])

        assert [r["file"] for r in rows] == [ok.name]
        assert rows[0]["cislo_faktury"] == "20260384" and rows[0]["variabilny_symbol"] == "20260384"
        assert stats["name_vs_qr_conflict"] == 1
