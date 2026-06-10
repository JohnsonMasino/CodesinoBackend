# Codesino_Backend

The official backend API for **Codesino Software Development Services** — a Django REST Framework-powered backend handling service requests, appointment bookings, newsletter subscriptions, and more.

---

## Tech Stack

- **Framework:** Django + Django REST Framework (DRF)
- **Database:** PostgreSQL
- **Deployment:** Render
- **Language:** Python 3.x

---

## Features

- 📬 **Newsletter Subscription** — Users can subscribe to the Codesino newsletter from the frontend
- 📅 **Appointment Booking** — Clients can book appointments via the frontend; requests appear in the Django admin panel for review
- 🛠️ **Service Requests** — Clients can submit requests for services directly from the frontend
- 🔐 **Django Admin Panel** — Admins can manage bookings, service requests, subscribers, and more
- 🌐 **REST API** — Clean and structured API endpoints consumed by the frontend

---

## Getting Started

### Prerequisites

- Python 3.10+
- PostgreSQL
- pip
- virtualenv

### Installation

1. **Clone the repository**

```bash
git clone https://github.com/Codesino-Sofware-Development-Services/Codesino_Backend.git
cd Codesino_Backend
```

2. **Create and activate a virtual environment**

```bash
python -m venv venv
source venv/bin/activate        # On Windows: venv\Scripts\activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

4. **Set up environment variables**

Create a `.env` file in the root directory and add the following:

```env
SECRET_KEY=your_django_secret_key
DEBUG=True
ALLOWED_HOSTS=localhost,127.0.0.1

# Database
DB_NAME=your_db_name
DB_USER=your_db_user
DB_PASSWORD=your_db_password
DB_HOST=localhost
DB_PORT=5432

# CORS (Frontend URL)
CORS_ALLOWED_ORIGINS=http://localhost:3000
```

5. **Apply migrations**

```bash
python manage.py migrate
```

6. **Create a superuser (for admin access)**

```bash
python manage.py createsuperuser
```

7. **Run the development server**

```bash
python manage.py runserver
```

The API will be available at `http://127.0.0.1:8000/`

---

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/newsletter/subscribe/` | Subscribe to the newsletter |
| POST | `/api/appointments/book/` | Book an appointment |
| POST | `/api/services/request/` | Submit a service request |
| GET | `/api/services/` | List available services |

> All appointment and service request submissions are accessible via the Django Admin at `/admin/`

---

## Deployment (Render)

This project is deployed on [Render](https://render.com).

### Environment Variables on Render

Set the following in your Render service dashboard under **Environment**:

```
SECRET_KEY
DEBUG=False
ALLOWED_HOSTS=your-render-domain.onrender.com
DATABASE_URL=your_postgres_connection_string
CORS_ALLOWED_ORIGINS=https://your-frontend-domain.com
```

### Build Command

```bash
pip install -r requirements.txt && python manage.py migrate
```

### Start Command

```bash
gunicorn your_project_name.wsgi:application
```

---

## Project Structure

```
Codesino_Backend/
├── appointments/        # Appointment booking app
├── newsletter/          # Newsletter subscription app
├── services/            # Service request app
├── core/                # Project settings and main URLs
├── manage.py
├── requirements.txt
└── .env
```

---

## Admin Panel

Access the Django admin dashboard at:

```
http://127.0.0.1:8000/admin/        # Local
https://your-domain.onrender.com/admin/   # Production
```

Admins can view and manage:
- Appointment booking requests
- Service requests
- Newsletter subscribers

---

## Contributing

This is a private company project. For internal contributions, please create a feature branch and open a pull request for review.

---

## License

© 2024 Codesino Software Development Services. All rights reserved.