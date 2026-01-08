# Category and Keyword Examples

This document provides example categories and keywords for different professions and use cases. Copy these into your Excel database's "Categories" sheet.

## General Professional

| Category | Keywords |
|----------|----------|
| Work | meeting, project, deadline, report, presentation, review, approval, task, assignment, team |
| Clients | client, customer, proposal, contract, agreement, consultation |
| Finance | invoice, payment, bill, receipt, transaction, expense, reimbursement, payroll |
| HR | benefits, vacation, pto, timeoff, insurance, onboarding, training |
| IT | support, ticket, password, access, vpn, system, maintenance, update |
| Marketing | campaign, analytics, social, content, seo, advertisement, promotion |
| Sales | lead, prospect, demo, pipeline, quota, opportunity, forecast |

## Healthcare Professional

| Category | Keywords |
|----------|----------|
| Patients | patient, appointment, consultation, examination, treatment, follow-up |
| Insurance | claim, authorization, coverage, copay, deductible, provider |
| Medical | prescription, lab, results, diagnosis, referral, imaging, radiology |
| Compliance | hipaa, compliance, regulation, audit, certification |
| Scheduling | schedule, booking, calendar, availability, reschedule |
| Pharmacy | medication, prescription, refill, pharmacy, dosage |

## Real Estate Professional

| Category | Keywords |
|----------|----------|
| Listings | listing, property, mls, showing, open house, virtual tour |
| Buyers | buyer, purchase, offer, pre-approval, mortgage, financing |
| Sellers | seller, cma, pricing, marketing, staging, photography |
| Transactions | closing, escrow, title, inspection, appraisal, contingency |
| Vendors | contractor, inspector, appraiser, photographer, stager |
| Legal | contract, addendum, disclosure, agreement, amendment |

## Education Professional

| Category | Keywords |
|----------|----------|
| Students | student, assignment, grade, homework, exam, quiz, attendance |
| Parents | parent, guardian, conference, meeting, permission, notification |
| Administration | faculty, staff, department, committee, policy, schedule |
| Curriculum | lesson, curriculum, standards, assessment, rubric, planning |
| Events | event, field trip, assembly, presentation, workshop, seminar |

## Legal Professional

| Category | Keywords |
|----------|----------|
| Cases | case, matter, client, docket, hearing, trial, deposition |
| Court | court, filing, motion, brief, pleading, subpoena, summons |
| Documents | contract, agreement, amendment, addendum, exhibit, affidavit |
| Billing | billing, invoice, retainer, fee, expense, time entry |
| Correspondence | correspondence, letter, notice, demand, response |
| Research | research, precedent, statute, regulation, case law |

## E-commerce / Retail

| Category | Keywords |
|----------|----------|
| Orders | order, purchase, transaction, checkout, cart, payment |
| Shipping | shipping, delivery, tracking, fulfillment, carrier, shipment |
| Returns | return, refund, exchange, rma, cancellation |
| Inventory | inventory, stock, sku, reorder, supplier, warehouse |
| Customers | customer, inquiry, question, feedback, complaint, review |
| Marketing | promotion, discount, sale, coupon, newsletter, campaign |

## Software Developer

| Category | Keywords |
|----------|----------|
| Code Review | review, pull request, pr, merge, commit, branch, code |
| Bugs | bug, issue, error, crash, fix, patch, debug |
| Deploy | deploy, deployment, release, production, staging, rollback |
| DevOps | server, infrastructure, cloud, aws, azure, docker, kubernetes |
| Meetings | standup, sprint, retro, retrospective, planning, scrum |
| Documentation | docs, documentation, readme, wiki, guide, tutorial |
| Security | security, vulnerability, cve, audit, penetration, breach |

## Freelancer / Consultant

| Category | Keywords |
|----------|----------|
| Projects | project, deliverable, milestone, scope, timeline, deadline |
| Proposals | proposal, quote, estimate, bid, statement of work, sow |
| Invoicing | invoice, payment, billing, rate, hours, expense |
| Clients | client, customer, stakeholder, meeting, call, check-in |
| Contracts | contract, agreement, terms, nda, non-disclosure |
| Marketing | portfolio, testimonial, referral, networking, lead |

## Small Business Owner

| Category | Keywords |
|----------|----------|
| Operations | operations, process, workflow, procedure, policy |
| Employees | employee, staff, hiring, payroll, benefits, performance |
| Vendors | vendor, supplier, wholesaler, distributor, manufacturer |
| Customers | customer, client, order, service, satisfaction, feedback |
| Finance | accounting, bookkeeping, taxes, profit, loss, revenue |
| Marketing | marketing, advertising, social media, website, seo, branding |
| Legal | legal, permit, license, compliance, insurance, liability |

## Content Creator

| Category | Keywords |
|----------|----------|
| Content | content, post, article, video, podcast, script, draft |
| Collaboration | collaboration, partnership, sponsor, brand, influencer |
| Analytics | analytics, metrics, views, engagement, subscribers, followers |
| Monetization | monetization, revenue, adsense, affiliate, sponsorship |
| Platform | youtube, instagram, tiktok, twitter, linkedin, facebook |
| Equipment | camera, microphone, lighting, editing, software, gear |

## Personal Use - Advanced

| Category | Keywords |
|----------|----------|
| Work | work, job, office, meeting, project, deadline, boss, colleague |
| Personal | personal, family, friend, birthday, anniversary, celebration |
| Finance | bank, credit, card, payment, bill, statement, tax, investment |
| Shopping | order, purchase, delivery, shipment, tracking, amazon, shop |
| Travel | flight, hotel, booking, reservation, trip, vacation, itinerary |
| Health | doctor, appointment, prescription, insurance, medical, health |
| Education | course, class, training, webinar, certificate, learning |
| Home | utilities, rent, mortgage, maintenance, repair, insurance |
| Entertainment | ticket, event, concert, movie, show, streaming, subscription |
| Newsletter | newsletter, digest, update, weekly, daily, roundup |
| Social | invitation, event, party, rsvp, gathering, meetup |
| Notifications | notification, alert, reminder, confirmation, verification |

## Tips for Creating Your Own Categories

### 1. Start with Email Patterns
Look at your last 50 emails and identify common themes.

### 2. Use Specific Keywords
- ✅ Good: "standup", "sprint", "pr review"
- ❌ Too broad: "email", "message", "hello"

### 3. Include Variations
```
meeting, mtg, call, zoom, teams, conference
invoice, inv, bill, billing, payment due
```

### 4. Use Sender Domains
```
Category: GitHub
Keyword: @github.com

Category: Team
Keyword: @yourcompany.com
```

### 5. Seasonal Categories
```
Category: Tax Season (Jan-Apr)
Keywords: tax, 1099, w2, deduction, accountant

Category: Holidays (Nov-Dec)
Keywords: holiday, gift, christmas, thanksgiving
```

### 6. Priority Levels
```
Category: Urgent
Keywords: urgent, asap, immediate, critical, emergency, important

Category: FYI
Keywords: fyi, info, informational, heads up, reminder
```

## How to Add These to Your Database

1. Open `EmailClusterDatabase.xlsx`
2. Go to the **Categories** sheet
3. For each keyword, add a row:

| Category | Keyword | Active | Created |
|----------|---------|--------|---------|
| Work | meeting | TRUE | 2026-01-08 10:00:00 |
| Work | project | TRUE | 2026-01-08 10:00:00 |
| Finance | invoice | TRUE | 2026-01-08 10:00:00 |

4. Save the file
5. Run the email clusterer again - it will use the new keywords!

## Testing Your Categories

After adding new categories:

```bash
# Test with a small batch
./email_clusterer.py --limit 10

# Check the EmailLogs sheet to see if categorization improved
open ~/Documents/EmailClusterDatabase.xlsx
```

## Refining Over Time

1. **Week 1**: Add basic categories (10-15 keywords)
2. **Week 2**: Review EmailLogs, add 10 more keywords
3. **Week 3**: Fine-tune by deactivating ineffective keywords
4. **Week 4**: Add advanced categories based on patterns
5. **Ongoing**: Add 2-3 keywords per week as you notice gaps

## Best Practices

- **Don't overdo it**: Start with 5-10 categories
- **Be specific**: Better to have precise keywords than too many generic ones
- **Review regularly**: Check EmailLogs monthly to see what's working
- **Seasonal adjustments**: Add/remove keywords based on time of year
- **Keep it simple**: If a category has 0 matches for a month, consider removing it

---

**Need help?** Check the main [README.md](README.md) for more guidance!
