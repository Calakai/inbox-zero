import { describe, expect, it } from "vitest";
import { extractMailFacts } from "./mail-facts";

describe("extractMailFacts", () => {
  it("pulls the balance, minimum payment and due date from a bank statement", () => {
    expect(
      extractMailFacts({
        subject: "Your October statement is ready",
        snippet: "Your statement balance is $1,284.17, due November 18.",
        body: "Hi Caleb, your October statement is ready. Statement balance: $1,284.17. Minimum payment: $35.00. Due date: November 18.",
      }),
    ).toEqual({ amounts: ["$1,284.17", "$35.00"], dueDate: "November 18" });
  });

  it("pulls the amount and pay-by date from a bill", () => {
    expect(
      extractMailFacts({
        subject: "Your electricity bill",
        snippet: "Amount due: $84.12",
        body: "Amount due: $84.12. Please pay by 11/05/2026 to avoid a late fee.",
      }),
    ).toEqual({ amounts: ["$84.12"], dueDate: "11/05/2026" });
  });

  it("lists a receipt's repeated total once, in order of first appearance", () => {
    expect(
      extractMailFacts({
        subject: "Order confirmation #8812",
        snippet: "Total $27.00",
        body: "Two bags of beans, $24.00. Shipping $3.00. Total $27.00.",
      }),
    ).toEqual({ amounts: ["$27.00", "$24.00", "$3.00"] });
  });

  it("leaves both fields out of a newsletter", () => {
    expect(
      extractMailFacts({
        subject: "UX Weekly #212: Motion that explains",
        snippet: "This week: motion, forms and 20% more reading.",
        body: "Issue #212, 2026. Join us at 2:00 for the live session.",
      }),
    ).toEqual({});
  });

  it.each([
    ["€12.50 off your next order", "€12.50"],
    ["Refund of £7.99 issued", "£7.99"],
    ["Deposit of CAD 40.00 received", "CAD 40.00"],
    ["You were charged 45.00 USD", "45.00 USD"],
    ["Paid C$19.99 today", "C$19.99"],
    ["Paid US$5 today", "US$5"],
    ["Rechnung über 12,50 € ist fällig", "12,50 €"],
    ["Order total ¥3,200", "¥3,200"],
  ])("keeps %s as written", (text, amount) => {
    expect(extractMailFacts({ subject: text }).amounts).toEqual([amount]);
  });

  it.each([
    "Order 112-4401 has shipped",
    "Re: issue #212",
    "Save 20% off everything",
    "See you in 2026",
    "Call starts at 2:00",
  ])("finds no amount in %s", (text) => {
    expect(extractMailFacts({ subject: text })).toEqual({});
  });

  it("keeps at most five amounts", () => {
    expect(extractMailFacts({ body: "$1 $2 $3 $4 $5 $6 $7" }).amounts).toEqual([
      "$1",
      "$2",
      "$3",
      "$4",
      "$5",
    ]);
  });

  it.each([
    ["Order #8812 $27.00", "$27.00"],
    ["Qty 2 $24.00", "$24.00"],
    ["Item 3 $9.99 each", "$9.99"],
    ["Call 1-800-555-0199 $5 fee", "$5"],
  ])("keeps the price after a number in %s", (text, amount) => {
    expect(extractMailFacts({ subject: text }).amounts).toEqual([amount]);
  });

  it.each([
    ["-$20.00 refund", "-$20.00"],
    ["\u2212$20.00 refund", "\u2212$20.00"],
    ["Paid CA$20 today", "CA$20"],
    ["Paid A$15 today", "A$15"],
    ["Paid AU$15.50 today", "AU$15.50"],
    ["Raised $5.5M in funding", "$5.5M"],
    ["Earn up to $10k a year", "$10k"],
  ])("keeps the sign, prefix and suffix in %s as written", (text, amount) => {
    expect(extractMailFacts({ subject: text }).amounts).toEqual([amount]);
  });

  it.each([
    ["due Mon 11/18", "Mon 11/18"],
    ["due 2026-11-18T00:00", "2026-11-18"],
  ])("reads the due date in %s with a weekday or time", (text, dueDate) => {
    expect(extractMailFacts({ body: text }).dueDate).toBe(dueDate);
  });

  it.each([
    ["Payment due: Friday, Nov 7, 2026.", "Friday, Nov 7, 2026"],
    ["Your card will be charged on Dec 1st.", "Dec 1st"],
    ["Payment date 2026-11-20", "2026-11-20"],
    ["The balance is due on 18 November 2026.", "18 November 2026"],
    ["Due by 11/5", "11/5"],
  ])("reads the due date in %s as written", (text, dueDate) => {
    expect(extractMailFacts({ body: text }).dueDate).toBe(dueDate);
  });

  it.each([
    "The delay was due to weather on November 3.",
    "Your report is due today.",
  ])("finds no due date in %s", (text) => {
    expect(extractMailFacts({ body: text }).dueDate).toBeUndefined();
  });

  it("takes the due date from the subject before the body", () => {
    expect(
      extractMailFacts({
        subject: "Payment due Nov 20",
        body: "Pay by December 1 to keep your discount.",
      }).dueDate,
    ).toBe("Nov 20");
  });
});
