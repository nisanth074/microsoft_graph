# microsoft_graph

Ruby gem for Microsoft Graph API

## Installation

Add the gem to your Rails app's Gemfile

```ruby
gem "microsoft_graph", git: "https://github.com/nisanth074/microsoft_graph", branch: "main"
```

and bundle install

```bash
bundle install
```

## Usage

```ruby
microsoft_graph = microsoft_graph.new(microsoft_oauth_access_token)

# Fetch messages received in the last 24 hours and print their subjects
one_day = 24 * 60 * 60
messages = microsoft_graph.messages_received_after(Time.now - one_day)
messages.each { |message| puts message.subject }

# Fetch the most recent message and print its subject
message = microsoft_graph.most_recent_message
messages.each { |message| puts message.subject }

# Find a message by its RFC 2822 Message-ID
message_id = "987747711.3069933.1624058987490.JavaMail.app-tm-001@idc-31-150"
microsoft_graph.find_message_by_message_id(message_id)

# Send a message
to = "john@example.com"
subject = "Account about to expire"
body = <<-MESSAGE_HTML
Hi John,

Just a friendly reminder that your account is about to expire in a week.

Best,
Joe
MESSAGE_HTML
message_hash = {
  "to" => [{
    "emailAddress" => {
      "address" => to
    }
  }],
  "subject" => subject,
  "body" => {
    "contentType" => "HTML",
    "content" => "html"
  },
}
message = microsoft_graph.send_message(message_hash)

# Send a reply to an existing message
message = microsoft_graph.most_recent_message
to = "john@example.com"
subject = "Re: Account about to expire"
body = <<-MESSAGE_HTML
Hi John,

Another friendly reminder that your account is about to expire in a day.

Best,
Joe
MESSAGE_HTML
reply_message_hash = {
  "to" => [{
    "emailAddress" => {
      "address" => to
    }
  }],
  "subject" => subject,
  "body" => {
    "contentType" => "HTML",
    "content" => "html"
  },
}
reply_message = message.send_reply(reply_message_hash)
```

## Resources

https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0

## Todos

- Add a usage example to fetch the most recent messages
- Add a usage example for attaching attachments
- Add a usage example for creating a draft message
- Port tests from the app
- Add a license
- Publish to rubygems.org
