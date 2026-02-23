"""Simple email formatting for AI-readable responses."""

from typing import List, Dict, Any
from datetime import datetime, timezone
from collections import defaultdict


def _sort_key_time(email: Dict[str, Any], reverse: bool = False) -> datetime:
    """Return a naive UTC datetime suitable for sorting.

    Handles the mix of offset-aware and offset-naive datetimes that
    Outlook COM can return.  Falls back to epoch-zero so emails
    without a timestamp sort to the end.
    """
    dt = email.get('received_time')
    if dt is None:
        return datetime.min if not reverse else datetime.max
    if hasattr(dt, 'tzinfo') and dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt


def _normalize_dt(dt: datetime) -> datetime:
    """Strip timezone to make any datetime naive-UTC for comparison."""
    if dt is None:
        return datetime.min
    if hasattr(dt, 'tzinfo') and dt.tzinfo is not None:
        return dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt

from ..config.config_reader import config


def format_mailbox_status(access_result: Dict[str, Any]) -> Dict[str, Any]:
    """Format mailbox access status for AI consumption."""
    
    return {
        "status": "success",
        "connection": {
            "outlook_connected": access_result.get("outlook_connected", False),
            "timestamp": datetime.now().isoformat()
        },
        "personal_mailbox": {
            "accessible": access_result.get("personal_accessible", False),
            "name": access_result.get("personal_name", "Personal Mailbox"),
            "retention_months": access_result.get("retention_personal_months", 6)
        },
        "shared_mailbox": {
            "configured": access_result.get("shared_configured", False),
            "accessible": access_result.get("shared_accessible", False),
            "name": access_result.get("shared_name", "Shared Mailbox"),
            "email": config.get("shared_mailbox_email", "Not configured"),
            "retention_months": access_result.get("retention_shared_months", 12)
        },
        "errors": access_result.get("errors", []),
        "notes": {
            "security_dialog": "You may need to grant permission when Outlook security dialog appears",
            "retention_info": "Search scope is limited by retention policy settings"
        }
    }


def format_email_chain(emails: List[Dict[str, Any]], search_subject: str) -> Dict[str, Any]:
    """Format email chain results for AI analysis."""
    
    if not emails:
        return {
            "status": "no_emails_found",
            "search_subject": search_subject,
            "message": f"No emails found for subject: '{search_subject}'"
        }
    
    # Group emails by conversation (based on subject similarity)
    conversations = group_by_conversation(emails)
    
    # Calculate statistics
    stats = {
        "total_emails": len(emails),
        "conversations": len(conversations),
        "date_range": get_date_range(emails),
        "mailbox_distribution": get_mailbox_distribution(emails),
        "participants": get_participants(emails)
    }
    
    # Format conversations chronologically
    formatted_conversations = []
    for conv_id, conv_emails in conversations.items():
        # Sort emails in conversation by time
        conv_emails.sort(key=lambda x: _sort_key_time(x))
        
        formatted_conv = {
            "conversation_id": conv_id,
            "email_count": len(conv_emails),
            "date_range": get_date_range(conv_emails),
            "participants": get_participants(conv_emails),
            "emails": [format_single_email(email) for email in conv_emails]
        }
        formatted_conversations.append(formatted_conv)
    
    # Sort conversations by most recent email
    formatted_conversations.sort(
        key=lambda x: max([parse_iso_time(e["received_time"]) for e in x["emails"] if e["received_time"]]),
        reverse=True
    )
    
    return {
        "status": "success",
        "search_subject": search_subject,
        "summary": stats,
        "conversations": formatted_conversations,
        "all_emails_chronological": [format_single_email(email) for email in sorted(emails, key=lambda x: _sort_key_time(x, reverse=True), reverse=True)]
    }


def format_alert_analysis(alerts: List[Dict[str, Any]], search_pattern: str) -> Dict[str, Any]:
    """Format alert analysis results for AI consumption."""
    
    if not alerts:
        return {
            "status": "no_alerts_found",
            "search_pattern": search_pattern,
            "message": f"No alerts found for pattern: '{search_pattern}'"
        }
    
    # Analyze alert patterns based on importance levels
    urgent_alerts = []
    normal_alerts = []
    
    analyze_importance = config.get_bool('analyze_importance_levels', True)
    
    for alert in alerts:
        # Check if alert is marked as high importance by sender
        is_urgent = alert.get('importance', 1) > 1
        
        # Additional urgency indicators if enabled
        if analyze_importance:
            subject = alert.get('subject', '').lower()
            # Simple urgency detection based on common urgent phrases
            urgent_phrases = ['urgent', 'critical', 'emergency', 'asap', 'immediate']
            is_urgent = is_urgent or any(phrase in subject for phrase in urgent_phrases)
        
        if is_urgent:
            urgent_alerts.append(alert)
        else:
            normal_alerts.append(alert)
    
    # Calculate alert frequency by day
    alert_frequency = calculate_daily_frequency(alerts)
    
    # Get recent alerts (last 10)
    recent_alerts = sorted(alerts, key=lambda x: _sort_key_time(x, reverse=True), reverse=True)[:10]
    
    # Summary statistics
    stats = {
        "total_alerts": len(alerts),
        "urgent_alerts": len(urgent_alerts),
        "normal_alerts": len(normal_alerts),
        "date_range": get_date_range(alerts),
        "mailbox_distribution": get_mailbox_distribution(alerts),
        "daily_frequency": alert_frequency,
        "response_indicators": analyze_responses(alerts)
    }
    
    return {
        "status": "success",
        "search_pattern": search_pattern,
        "summary": stats,
        "urgent_alerts": [format_single_email(alert) for alert in urgent_alerts[:5]],  # Top 5 urgent
        "recent_alerts": [format_single_email(alert) for alert in recent_alerts],
        "timeline": create_alert_timeline(alerts),
        "recommendations": generate_alert_recommendations(stats, urgent_alerts)
    }


def format_single_email(email: Dict[str, Any]) -> Dict[str, Any]:
    """Format a single email for AI consumption."""

    formatted = {
        "subject": email.get('subject', 'No Subject'),
        "sender_name": email.get('sender_name', 'Unknown'),
        "sender_email": email.get('sender_email', ''),
        "to": email.get('to_recipients', []),
        "cc": email.get('cc_recipients', []),
        "recipients": email.get('recipients', []),
        "folder": email.get('folder_name', 'Unknown'),
        "mailbox": email.get('mailbox_type', 'unknown'),
        "body_preview": email.get('body', '')[:500] if email.get('body') else '',
        "attachments": email.get('attachments_count', 0),
        "importance": get_importance_text(email.get('importance', 1)),
        "unread": email.get('unread', False),
        "size_kb": round(email.get('size', 0) / 1024, 1)
    }

    # Add timestamp if configured
    if config.get_bool('include_timestamps', True):
        received_time = email.get('received_time')
        formatted["received_time"] = received_time.isoformat() if received_time else None

    return formatted


def group_by_conversation(emails: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """Group emails by conversation based on subject similarity."""
    conversations = defaultdict(list)
    
    for email in emails:
        subject = email.get('subject', '').strip()
        
        # Clean subject for grouping (remove Re:, Fwd:, etc.)
        clean_subject = subject
        prefixes = ['re:', 'fwd:', 'fw:', 'reply:', 'forward:']
        for prefix in prefixes:
            if clean_subject.lower().startswith(prefix):
                clean_subject = clean_subject[len(prefix):].strip()
        
        # Use cleaned subject as conversation key
        conv_key = clean_subject.lower()
        conversations[conv_key].append(email)
    
    return dict(conversations)


def get_date_range(emails: List[Dict[str, Any]]) -> Dict[str, str]:
    """Get date range of emails."""
    if not emails:
        return {"first": None, "last": None}
    
    dates = [email.get('received_time') for email in emails if email.get('received_time')]
    if not dates:
        return {"first": None, "last": None}

    normalized = [_normalize_dt(d) for d in dates]
    return {
        "first": min(normalized).isoformat(),
        "last": max(normalized).isoformat()
    }


def get_mailbox_distribution(emails: List[Dict[str, Any]]) -> Dict[str, int]:
    """Get distribution of emails across mailboxes."""
    distribution = {"personal": 0, "shared": 0, "unknown": 0}
    
    for email in emails:
        mailbox_type = email.get('mailbox_type', 'unknown')
        if mailbox_type in distribution:
            distribution[mailbox_type] += 1
        else:
            distribution['unknown'] += 1
    
    return distribution


def get_participants(emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Get list of email participants with counts and role breakdown.

    Tracks how many times each person appeared as sender, To, or CC
    so you can see who is central to a thread.
    """
    # keyed by lowercase email (or name if no email)
    stats = {}  # key -> {name, email, sent, to, cc}

    def _touch(key: str, name: str, email: str, role: str):
        if key not in stats:
            stats[key] = {'name': name, 'email': email, 'sent': 0, 'to': 0, 'cc': 0}
        stats[key][role] += 1
        # prefer the richer value
        if email and not stats[key]['email']:
            stats[key]['email'] = email
        if name and (not stats[key]['name'] or stats[key]['name'] == 'Unknown'):
            stats[key]['name'] = name

    for email in emails:
        sender_name = email.get('sender_name', 'Unknown')
        sender_email = email.get('sender_email', '')
        key = (sender_email or sender_name).lower()
        _touch(key, sender_name, sender_email, 'sent')

        for entry in email.get('to_recipients', []):
            e = entry.get('email', '') if isinstance(entry, dict) else ''
            n = entry.get('name', '') if isinstance(entry, dict) else str(entry)
            k = (e or n).lower()
            if k:
                _touch(k, n, e, 'to')

        for entry in email.get('cc_recipients', []):
            e = entry.get('email', '') if isinstance(entry, dict) else ''
            n = entry.get('name', '') if isinstance(entry, dict) else str(entry)
            k = (e or n).lower()
            if k:
                _touch(k, n, e, 'cc')

    # Sort by total participation
    participants = []
    for info in sorted(stats.values(),
                       key=lambda x: x['sent'] + x['to'] + x['cc'],
                       reverse=True):
        total = info['sent'] + info['to'] + info['cc']
        participants.append({
            "name": info['name'],
            "email": info['email'],
            "participation_count": total,
            "sent": info['sent'],
            "to_count": info['to'],
            "cc_count": info['cc'],
        })

    return participants[:20]  # Top 20 participants


def calculate_daily_frequency(alerts: List[Dict[str, Any]]) -> float:
    """Calculate average alerts per day."""
    if not alerts:
        return 0.0
    
    dates = [alert.get('received_time').date() for alert in alerts if alert.get('received_time')]
    if not dates:
        return 0.0
    
    date_range_days = (max(dates) - min(dates)).days + 1
    return round(len(alerts) / max(date_range_days, 1), 2)


def analyze_responses(alerts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Analyze response patterns in alerts."""
    replies = sum(1 for alert in alerts if alert.get('subject', '').lower().startswith(('re:', 'reply:')))
    total = len(alerts)
    
    return {
        "replies_found": replies,
        "total_alerts": total,
        "response_rate_percent": round((replies / total) * 100, 1) if total > 0 else 0
    }


def create_alert_timeline(alerts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Create chronological timeline of alerts."""
    timeline = []
    
    # Sort alerts by time
    sorted_alerts = sorted(alerts, key=lambda x: _sort_key_time(x))
    
    for alert in sorted_alerts:
        timeline_entry = {
            "timestamp": alert.get('received_time').isoformat() if alert.get('received_time') else None,
            "subject": alert.get('subject', 'No Subject')[:100],  # Truncate long subjects
            "sender": alert.get('sender_name', 'Unknown'),
            "mailbox": alert.get('mailbox_type', 'unknown'),
            "folder": alert.get('folder_name', 'Unknown'),
            "importance": get_importance_text(alert.get('importance', 1))
        }
        timeline.append(timeline_entry)
    
    return timeline


def generate_alert_recommendations(stats: Dict[str, Any], urgent_alerts: List[Dict[str, Any]]) -> List[str]:
    """Generate actionable recommendations based on alert analysis."""
    recommendations = []
    
    total_alerts = stats.get('total_alerts', 0)
    urgent_count = stats.get('urgent_alerts', 0)
    daily_frequency = stats.get('daily_frequency', 0)
    
    # High frequency alerts
    if daily_frequency > 5:
        recommendations.append(f"High alert frequency detected ({daily_frequency} alerts/day) - investigate root causes")
    
    # High urgency rate
    if urgent_count > 0 and total_alerts > 0:
        urgent_rate = (urgent_count / total_alerts) * 100
        if urgent_rate > 30:
            recommendations.append(f"High urgency rate ({urgent_rate:.1f}%) - review alert thresholds")
    
    # Response rate analysis
    response_rate = stats.get('response_indicators', {}).get('response_rate_percent', 0)
    if response_rate < 50:
        recommendations.append(f"Low response rate ({response_rate}%) - review alert response procedures")
    
    # Recent urgent alerts
    if urgent_count > 0:
        recommendations.append(f"Review {urgent_count} urgent alerts for immediate action")
    
    # Mailbox distribution
    mailbox_dist = stats.get('mailbox_distribution', {})
    if mailbox_dist.get('personal', 0) > 0 and mailbox_dist.get('shared', 0) == 0:
        recommendations.append("Alerts found only in personal mailbox - verify shared mailbox routing")
    
    if not recommendations:
        recommendations.append("No immediate issues detected - continue monitoring")
    
    return recommendations


def get_importance_text(importance: int) -> str:
    """Convert importance number to text."""
    importance_map = {0: "Low", 1: "Normal", 2: "High"}
    return importance_map.get(importance, "Normal")


def parse_iso_time(iso_string: str) -> datetime:
    """Parse ISO timestamp string to a naive-UTC datetime for sorting."""
    try:
        dt = datetime.fromisoformat(iso_string.replace('Z', '+00:00'))
        return _normalize_dt(dt)
    except (ValueError, AttributeError):
        return datetime.min
